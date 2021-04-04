VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm应付款查询 
   BackColor       =   &H8000000A&
   Caption         =   "药品应付款查询"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frm应付款查询.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   635
      SimpleText      =   $"frm应付款查询.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm应付款查询.frx":0351
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9340
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
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   345
      Left            =   4710
      TabIndex        =   8
      Top             =   2880
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   609
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "付款明细帐"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "已付清单"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "未付清单"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   1005
      Left            =   3330
      TabIndex        =   7
      Top             =   1290
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1773
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2550
      Width           =   3000
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2190
      Top             =   2820
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
            Picture         =   "frm应付款查询.frx":0BE5
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":0D3F
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":0E99
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   5640
      Top             =   60
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
            Picture         =   "frm应付款查询.frx":11B3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":13CF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":15EB
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":1805
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":1A1F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":1C3B
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   4920
      Top             =   60
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
            Picture         =   "frm应付款查询.frx":1E57
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":2073
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":228F
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":24A9
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":26C3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款查询.frx":28DF
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8160
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
               Object.ToolTipText     =   "重置条件"
               Object.Tag             =   "重置"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Find"
               Description     =   "定位单据"
               Object.ToolTipText     =   "定位"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1215
      Left            =   2910
      TabIndex        =   9
      Top             =   3240
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "明细信息"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3255
      TabIndex        =   10
      Top             =   2610
      Width           =   3015
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "汇总信息"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3450
      TabIndex        =   4
      Top             =   930
      Width           =   1950
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
      Begin VB.Menu mnuViewUnit 
         Caption         =   "售价单位(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "门诊单位(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "住院单位(&Z)"
         Index           =   2
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "药库单位(&K)"
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "条件重置(&J)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "单据定位(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
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
Attribute VB_Name = "frm应付款查询"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnLoad As Boolean
Dim mdatBegin As Date, mdatEnd As Date            '查询的时间范围
Dim mstrData As String
Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mstrKey As String       '前一个树节点的关键值
Dim mlngID As Long          '前一个药品供应商的ID

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillTree
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim i As Double
    mblnLoad = True
    RestoreWinState Me, App.ProductName
    
    i = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewUnit", "0"))
    If i < 0 Or i > 3 Then
        i = 0
    Else
        i = Int(i)
    End If
    mnuViewUnit(i).Checked = True
    
    If glngSys \ 100 = 8 Then
        '药店系统
        mnuViewUnit(1).Visible = False
        mnuViewUnit(2).Visible = False
        mnuViewUnit(3).Caption = "采购单位(&K)"
    End If
    '得到查询的时间范围
    mdatEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    mdatBegin = DateAdd("m", -1, mdatEnd) + 1
    mstrData = "0000"
    Call InitSum
End Sub

Private Sub InitSum()
'初始化汇总表的样式
    With mshSum
        ClearGrid mshSum, 5
        .TextMatrix(0, 0) = "药品供应商"
        .TextMatrix(0, 1) = "期初应付"
        .TextMatrix(0, 2) = "本期赊购"
        .TextMatrix(0, 3) = "本期支付"
        .TextMatrix(0, 4) = "期末应付"
        
        .colWidth(0) = 2000
        .colWidth(1) = 1500
        .colWidth(2) = 1500
        .colWidth(3) = 1500
        .colWidth(4) = 1500
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    Dim i As Integer
    
    For i = 0 To 3
        If mnuViewUnit(i).Checked = True Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewUnit", i
        End If
    Next
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
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
    '左边
    lblSum.Top = sngTop
    lblSum.Left = picV.Left + picV.Width
    If ScaleWidth - lblSum.Left > 0 Then lblSum.Width = ScaleWidth - lblSum.Left
    
    mshSum.Left = lblSum.Left
    picH.Left = lblSum.Left
    lblDetail.Left = lblSum.Left
    tabMain.Left = lblDetail.Left + lblDetail.Width + 60
    mshDetail.Left = lblSum.Left
    
    mshSum.Width = lblSum.Width
    picH.Width = lblSum.Width
    tabMain.Width = ScaleWidth - tabMain.Left
    mshDetail.Width = lblSum.Width
    
    'mshSum的位置
    mshSum.Top = lblSum.Top + lblSum.Height
    'picH的位置
    picH.Top = mshSum.Top + mshSum.Height
    'tabMain的位置
    lblDetail.Top = picH.Top + picH.Height + 15
    tabMain.Top = picH.Top + picH.Height
    'mshDetail的位置
    mshDetail.Top = tabMain.Top + tabMain.Height
    mshDetail.Height = IIf(sngBottom - mshDetail.Top > 0, sngBottom - mshDetail.Top, 0)
    
    Refresh
End Sub

Private Sub mnuViewFind_Click()
'按药品供应商与单据号定位
    Dim str单据号 As String, str供应商ID As String
    Dim rsTemp As New ADODB.Recordset
    Dim nod As MSComctlLib.node, lngRow As Long, lngCol As Long
    
    If frm应付款定位.Get定位条件(str单据号, str供应商ID) = False Then
        Exit Sub
    End If
    
    If str单据号 <> "" Then
        '根据单据号找到供应商
        gstrSQL = "select 供药单位ID from 药品收发记录 where NO='" & str单据号 & "' and 单据=1 and 供药单位ID is not null"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = True Then
            MsgBox "单据号为 " & str单据号 & " 外购入库单没有找到。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str供应商ID = rsTemp("供药单位ID")
        rsTemp.Close
    End If
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes("C" & str供应商ID)
    If Err <> 0 Then
        MsgBox "没有发现指定供应商，可能已经被停用。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    nod.Selected = True
    nod.EnsureVisible
    Call FillSum
    
    If str单据号 <> "" Then
        '找到单据所在列
        If tabMain.SelectedItem.Index = 1 Then
            lngCol = 1
        Else
            lngCol = 0
        End If
        
        With mshDetail
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, lngCol) = str单据号 Then
                    .TopRow = lngRow
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub mnuViewOpen_Click()
    If frmTimeSet.GetTimeScope(mdatBegin, mdatEnd, mstrData, Me) = True Then
        mstrKey = ""
        Call FillSum
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FillTree
End Sub

Private Sub mnuViewUnit_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewUnit(i).Checked = False
    Next
    mnuViewUnit(Index).Checked = True
    
    Call FillDetail
End Sub

Private Sub mshSum_EnterCell()
    If mlngID = mshSum.RowData(mshSum.Row) Then Exit Sub
    mlngID = mshSum.RowData(mshSum.Row)
    Call FillDetail
End Sub

Private Sub mshSum_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal node As MSComctlLib.node)
    FillSum
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
        If sngTemp > 500 And ScaleWidth - (sngTemp + picV.Width) > 600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub picH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picH.Top + y - msngStartY
        If sngTemp > mshSum.Top + 600 And ScaleHeight - (sngTemp + picH.Height) > 1600 Then
            picH.Top = sngTemp
            mshSum.Height = picH.Top - mshSum.Top
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


Private Sub tabMain_Click()
    Call FillDetail
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            mnuViewOpen_Click
        Case "Find"
            mnuViewFind_Click
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
   Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
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
    
    If mshSum Is ActiveControl Then
        Set objPrint.Body = mshSum
        objPrint.Title.Text = "应付购药款汇总信息"
        objRow.Add " "
        objRow.Add "查询时间：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.用户姓名
        objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = mshDetail
        objPrint.Title.Text = tabMain.SelectedItem.Caption
        objRow.Add "药品供应商：" & lblDetail.Caption
        objRow.Add "查询时间：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.用户姓名
        objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
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

Private Function FillTree() As Boolean
'功能:装入药品供应商
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strkey As String
    
    mstrKey = ""     '全面刷新时就相当于用户没点过任何节点
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strkey = tvwMain_S.SelectedItem.Key
    End If
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select id,上级id,编码,名称,末级 from 药品供应商  where 撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null " & _
         " start with 上级ID is null  connect by prior id=上级ID "
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "“药品供应商”的信息不全，无法进行查询。", vbExclamation, gstrSysName
        FillTree = False
        Exit Function
    End If
    
    
    With tvwMain_S.Nodes
        .Clear
        .Add , , "Root", "所有药品供应商", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        Do Until rsTemp.EOF
            '得出正确的图标
            strTemp = IIf(rsTemp("末级") = 1, "Item", "Class")
            '添加节点
            If IsNull(rsTemp("上级id")) Then
                .Add "Root", tvwChild, "C" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), strTemp, strTemp
            Else
                .Add "C" & rsTemp("上级id"), tvwChild, "C" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), strTemp, strTemp
            End If
            tvwMain_S.Nodes("C" & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    Dim nod As node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strkey)
    If Err <> 0 Then
        Err.Clear
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
    Else
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
    End If
    Call FillSum
    FillTree = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillSum()
'功能:装入各种统计数据
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 4) As Double
    Dim lngRow As Long
    Dim blnSum As Boolean        '合计的显示
    
    
    stbThis.Panels(2).Text = "时间范围：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")

    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    If mstrKey = tvwMain_S.SelectedItem.Key Then Exit Sub
    mstrKey = tvwMain_S.SelectedItem.Key
    '开始查询
    
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd, "yyyyMMdd")
    MousePointer = 11
    '首先得到子查询的SQL语句
    If tvwMain_S.SelectedItem.Image = "Item" Then
        gstrSQL = " and A.单位ID=" & Mid(mstrKey, 2)
    ElseIf tvwMain_S.SelectedItem.Image = "Root" Then
        gstrSQL = " and A.单位ID in (select ID from 药品供应商 start with 上级ID is null connect by prior id=上级ID)"
    Else
        gstrSQL = " and A.单位ID in (select ID from 药品供应商 start with 上级ID =" & Mid(mstrKey, 2) & " connect by prior id=上级ID)"
    End If
    If Mid(mstrData, 1, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.期初应付<>0 "
    End If
    If Mid(mstrData, 2, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.本期赊购<>0 "
    End If
    If Mid(mstrData, 3, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.本期支付<>0 "
    End If
    If Mid(mstrData, 4, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.期末应付<>0 "
    End If
    
    '再得到完整的SQL语句
    gstrSQL = "select B.名称,B.ID,A.期初应付,A.本期赊购,A.本期支付,A.期末应付 from " & _
            "(select 单位ID,sum(余额-期初应付+期初付款) as 期初应付,sum(期初应付-期末应付) as 本期赊购 " & _
            "            ,sum(期初付款-期末付款) as 本期支付,sum(余额-期末应付+期末付款) as 期末应付 " & _
            "from( " & _
            "select 单位ID,金额 as 期初付款, " & _
            "    decode(sign(to_char(审核日期,'yyyymmdd')-'" & strEnd & "'),1,金额,0) as 期末付款, " & _
            "    0 as 期初应付,0 as 期末应付,0 as 余额 from 药品付款记录 " & _
            "    where 审核日期>=to_date('" & strBegin & "','yyyyMMdd') " & _
            "Union All " & _
            "select A.供药单位ID 单位ID,0 as 期初付款,0 as 期末付款, " & _
            "    A.发票金额 as 期初应付,decode(sign(to_char(B.审核日期,'yyyymmdd')-'" & strEnd & "'),1,A.发票金额,0) as 期末应付,0 as 余额 from 药品应付记录 A,药品收发记录 B " & _
            "    where A.收发ID=B.ID and B.审核日期>=to_date('" & strBegin & "','yyyyMMdd') " & _
            "Union All " & _
            "select 供药商ID 单位ID,0 as 期初付款,0 as 期末付款,0 as 期初应付,0 as 期末应付,金额 as 余额 from 药品应付余额 " & _
            "    where 性质=1) " & _
            "group by 单位ID)A,药品供应商 B " & _
            "where A.单位ID=B.ID  " & gstrSQL
    Call OpenRecordset(rsTemp, Me.Caption)
    
    mshSum.Redraw = False
    If rsTemp.RecordCount = 0 Then
        ClearGrid mshSum
    Else
        If rsTemp.RecordCount = 1 Then
            '只有一行，就不显示合计了
            mshSum.Rows = 2
            blnSum = False
        Else
            mshSum.Rows = rsTemp.RecordCount + 2
            blnSum = True
        End If
    End If
    lngRow = 1
    With mshSum
        Do Until rsTemp.EOF
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, 0) = rsTemp("名称")
            .TextMatrix(lngRow, 1) = Format(rsTemp("期初应付"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("本期赊购"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 3) = Format(rsTemp("本期支付"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 4) = Format(rsTemp("期末应付"), "###########0.00;-###########0.00; ; ")
            If blnSum = True Then
                dblSum(1) = dblSum(1) + rsTemp("期初应付")
                dblSum(2) = dblSum(2) + rsTemp("本期赊购")
                dblSum(3) = dblSum(3) + rsTemp("本期支付")
                dblSum(4) = dblSum(4) + rsTemp("期末应付")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If blnSum = True Then
            .TextMatrix(lngRow, 0) = "  合计"
            .TextMatrix(lngRow, 1) = Format(dblSum(1), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(dblSum(2), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 3) = Format(dblSum(3), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 4) = Format(dblSum(4), "###########0.00;-###########0.00; ; ")
        End If
    End With
    mshSum.Redraw = True
    
    MousePointer = 0
    Call FillDetail
End Sub

Private Sub FillDetail()
'功能:装入各种明细数据
    
    MousePointer = 11
    If mshSum.RowData(mshSum.Row) <> 0 Then
        lblDetail.Caption = mshSum.TextMatrix(mshSum.Row, 0)
    Else
        lblDetail.Caption = "明细信息"
    End If
    lblDetail.ToolTipText = lblDetail.Caption
    
    mshDetail.Redraw = False
    Select Case tabMain.SelectedItem.Index
        Case 2
            Call Fill已付清单
        Case 3
            Call Fill未付清单
        Case Else
            Call Fill明细帐
    End Select
    mshDetail.Redraw = True
    MousePointer = 0
    
    Call MenuSet
End Sub

Private Sub Fill明细帐()
'功能:装入明细帐数据
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 2) As Double, dblBalance As Double
    Dim lngRow As Long, lngID As Long
    
    '初始化表格
    ClearGrid mshDetail, 11
    
    With mshDetail
        .MergeCells = flexMergeNever
        .TextMatrix(0, 0) = "日期":     .colWidth(0) = 1100: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "单据号":   .colWidth(1) = 1000: .ColAlignment(1) = 1
        .TextMatrix(0, 2) = "摘要":     .colWidth(2) = 1750: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "单位":     .colWidth(3) = 600: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "批号":     .colWidth(4) = 900: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "失效期":     .colWidth(5) = 1100: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "采购数量": .colWidth(6) = 900: .ColAlignment(6) = 7
        .TextMatrix(0, 7) = "采购价":   .colWidth(7) = 900: .ColAlignment(7) = 7
        .TextMatrix(0, 8) = "应付金额": .colWidth(8) = 1000: .ColAlignment(8) = 7
        .TextMatrix(0, 9) = "已付金额": .colWidth(9) = 1000: .ColAlignment(9) = 7
        .TextMatrix(0, 10) = "余额":     .colWidth(10) = 1000: .ColAlignment(10) = 7
    End With
    '得到查询条件
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '开始查询
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    '首先得到期末余额
    gstrSQL = "select sum(金额) as 余额 from( " & _
            "select 金额  from 药品付款记录 " & _
            "    where 审核日期>=to_date('" & strEnd & "','yyyyMMdd')  and 单位ID=" & lngID & _
            " Union All " & _
            "select -1 * A.发票金额 as 金额 from 药品应付记录 A,药品收发记录 B " & _
            "    where A.收发ID=B.ID and B.审核日期>=to_date('" & strEnd & "','yyyyMMdd')  and A.供药单位ID=" & lngID & _
            " Union All " & _
            "select 金额  from 药品应付余额 " & _
            "    where 性质=1 and 供药商ID=" & lngID & ") "
    Call OpenRecordset(rsTemp, Me.Caption)
    
    dblBalance = IIf(IsNull(rsTemp("余额")), 0, rsTemp("余额"))
    rsTemp.Close
    
    
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",D.门诊单位 as 单位,B.实际数量/D.门诊包装 as 采购数量,B.成本价*D.门诊包装 as 采购价"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",D.住院单位 as 单位,B.实际数量/D.住院包装 as 采购数量,B.成本价*D.住院包装 as 采购价"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",D.药库单位 as 单位,B.实际数量/D.药库包装 as 采购数量,B.成本价*D.药库包装 as 采购价"
    Else
        gstrSQL = ",D.售价单位 as 单位,B.实际数量 as 采购数量,B.成本价 as 采购价"
    End If
    '再得到明细帐
    gstrSQL = " select * from( " & _
              "  select to_char(审核日期,'yyyy-MM-dd') as 日期,'付'||NO as NO,序号, " & _
              "       decode(预付款,1,'预付款',decode(记录状态,2,'预付款',摘要))||'('||结算方式||')' as 摘要, " & _
              "       '' as 批号,'' as 效期,'' as 单位,0 as 采购数量,0 as 采购价,0 as 应付金额,金额 as 已付金额 " & _
              "       From 药品付款记录 " & _
              "       where 审核日期>=to_date('" & strBegin & "','yyyyMMdd') and 审核日期<to_date('" & strEnd & "','yyyyMMdd') and 单位ID=" & lngID & _
              "  Union All " & _
              "  select to_char(B.审核日期,'yyyy-MM-dd') as 日期,B.NO,B.序号,C.通用名称||'('||D.规格||')' as 摘要, " & _
              "       批号,to_char(效期,'yyyy-MM-dd') as 效期" & gstrSQL & ",A.发票金额 as 应付金额,0 as 已付金额 " & _
              "       from 药品应付记录 A,药品收发记录 B,药品目录 D,药品信息 C " & _
              "       where B.审核日期>=to_date('" & strBegin & "','yyyyMMdd') and B.审核日期<to_date('" & strEnd & "','yyyyMMdd') and A.供药单位ID=" & lngID & _
              "             and A.收发ID = B.ID And B.药品ID=D.药品ID and C.药名ID=D.药名ID ) " & _
              "  order by 日期,no,序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    mshDetail.Rows = rsTemp.RecordCount + 3
    lngRow = 2
    With mshDetail
        .TextMatrix(1, 0) = Format(mdatBegin, "yyyy-MM-dd")
        .TextMatrix(1, 2) = "期初余额"
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = rsTemp("日期")
            .TextMatrix(lngRow, 1) = rsTemp("NO")
            .TextMatrix(lngRow, 2) = IIf(IsNull(rsTemp("摘要")), "", rsTemp("摘要"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("单位")), "", rsTemp("单位"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("批号")), "", rsTemp("批号"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("效期")), "", rsTemp("效期"))
            .TextMatrix(lngRow, 6) = Format(rsTemp("采购数量"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 7) = Format(rsTemp("采购价"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 8) = Format(rsTemp("应付金额"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(rsTemp("已付金额"), "###########0.00;-###########0.00; ; ")
                
            dblSum(1) = dblSum(1) + rsTemp("应付金额")
            dblSum(2) = dblSum(2) + rsTemp("已付金额")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .TextMatrix(lngRow, 0) = Format(mdatEnd, "yyyy-MM-dd")
        .TextMatrix(lngRow, 2) = "合计"
        .TextMatrix(lngRow, 8) = Format(dblSum(1), "###########0.00;-###########0.00; ; ")
        .TextMatrix(lngRow, 9) = Format(dblSum(2), "###########0.00;-###########0.00; ; ")
        .TextMatrix(lngRow, 10) = Format(dblBalance, "###########0.00;-###########0.00; ; ")
        
        
        Do Until lngRow = 1
            lngRow = lngRow - 1
            .TextMatrix(lngRow, 10) = Format(dblBalance, "###########0.00;-###########0.00; ; ")
            dblBalance = dblBalance + Val(.TextMatrix(lngRow, 9)) - Val(.TextMatrix(lngRow, 8))
        Loop
    End With

End Sub

Private Sub Fill已付清单()
'功能:装入已付清单
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    
    '初始化表格
    ClearGrid mshDetail, 11
    With mshDetail
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .TextMatrix(0, 0) = "入库单据号":   .colWidth(0) = 1000: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "单据金额":     .colWidth(1) = 1000: .ColAlignment(1) = 7
        .TextMatrix(0, 2) = "付款单据号":   .colWidth(2) = 1000: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "日期":         .colWidth(3) = 1100: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "药品名称":     .colWidth(4) = 1500: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "规格":         .colWidth(5) = 900: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "单位":         .colWidth(6) = 900: .ColAlignment(6) = 1
        .TextMatrix(0, 7) = "批号":         .colWidth(7) = 900: .ColAlignment(7) = 1
        .TextMatrix(0, 8) = "失效期":       .colWidth(8) = 1100: .ColAlignment(8) = 1
        .TextMatrix(0, 9) = "数量":         .colWidth(9) = 1000: .ColAlignment(9) = 7
        .TextMatrix(0, 10) = "金额":        .colWidth(10) = 1100: .ColAlignment(10) = 7
    End With
    '得到查询条件
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '开始查询
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",C.门诊单位 as 单位,B.实际数量/C.门诊包装 as 数量"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",C.住院单位 as 单位,B.实际数量/C.住院包装 as 数量"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",C.药库单位 as 单位,B.实际数量/C.药库包装 as 数量"
    Else
        gstrSQL = ",C.售价单位 as 单位,B.实际数量 as 数量"
    End If
    '得到已付清单
    gstrSQL = "select B.NO as 入库单据号,b.序号,E.NO as 付款单据号,to_char(E.审核日期,'yyyy-MM-dd') as 日期, " & _
              "         D.通用名称 as 名称,C.规格,B.批号,to_char(B.效期,'yyyy-MM-dd') as 效期" & gstrSQL & ",A.发票金额 as 金额 " & _
              "    from 药品应付记录 A,药品收发记录 B,药品目录 C,药品信息 D ,药品付款记录 E " & _
              "    Where A.收发ID = B.ID And B.药品ID = C.药品ID And C.药名ID = D.药名ID And A.付款序号 = e.付款序号 And e.序号 = 1 " & _
              "        and E.审核日期>=to_date('" & strBegin & "','yyyyMMdd') and E.审核日期<to_date('" & strEnd & "','yyyyMMdd') and A.供药单位ID=" & lngID & _
              " and E.单位ID= " & lngID & "  and E.记录状态<>2 order by B.NO,B.序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        mshDetail.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With mshDetail
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("入库单据号")), " ", rsTemp("入库单据号"))
            .TextMatrix(lngRow, 2) = rsTemp("付款单据号")
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("日期")), " ", rsTemp("日期"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("规格")), "", rsTemp("规格"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rsTemp("单位")), "", rsTemp("单位"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rsTemp("批号")), "", rsTemp("批号"))
            .TextMatrix(lngRow, 8) = IIf(IsNull(rsTemp("效期")), "", rsTemp("效期"))
            .TextMatrix(lngRow, 9) = Format(rsTemp("数量"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 10) = Format(rsTemp("金额"), "###########0.00;-###########0.00; ; ")
                
            dblSum = dblSum + rsTemp("金额")
            lngRow = lngRow + 1
            
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, 0) = "合计"
            .TextMatrix(lngRow, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 10) = Format(dblSum, "###########0.00;-###########0.00; ; ")
        End If
        '再算单据的合计金额
        lngRow = 1
        Do While lngRow < .Rows - 1
            dblSum = Val(.TextMatrix(lngRow, 10))
            lngTemp = lngRow + 1
            
            Do While lngTemp < .Rows - 1
                If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) Then
                    dblSum = dblSum + Val(.TextMatrix(lngTemp, 10))
                Else
                    Exit Do
                End If
                lngTemp = lngTemp + 1
            Loop
            For lngCount = lngRow To lngTemp - 1
                .TextMatrix(lngCount, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            Next
            lngRow = lngTemp
        Loop
    End With
End Sub

Private Sub Fill未付清单()
'功能:装入已付清单
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    
    '初始化表格
    ClearGrid mshDetail, 10
    With mshDetail
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .TextMatrix(0, 0) = "入库单据号":   .colWidth(0) = 1000: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "单据金额":     .colWidth(1) = 1000: .ColAlignment(1) = 7
        .TextMatrix(0, 2) = "日期":         .colWidth(2) = 1100: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "药品名称":     .colWidth(3) = 1500: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "规格":         .colWidth(4) = 900: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "单位":         .colWidth(5) = 900: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "批号":         .colWidth(6) = 900: .ColAlignment(6) = 1
        .TextMatrix(0, 7) = "失效期":         .colWidth(7) = 1100: .ColAlignment(7) = 1
        .TextMatrix(0, 8) = "数量":         .colWidth(8) = 1000: .ColAlignment(8) = 7
        .TextMatrix(0, 9) = "金额":         .colWidth(9) = 1100: .ColAlignment(9) = 7
    End With
    '得到查询条件
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '开始查询
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    
    '得到未付清单
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",C.门诊单位 as 单位,B.实际数量/C.门诊包装 as 数量"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",C.住院单位 as 单位,B.实际数量/C.住院包装 as 数量"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",C.药库单位 as 单位,B.实际数量/C.药库包装 as 数量"
    Else
        gstrSQL = ",C.售价单位 as 单位,B.实际数量 as 数量"
    End If
    gstrSQL = "select B.NO as 入库单据号,b.序号,to_char(B.审核日期,'yyyy-MM-dd') as 日期, " & _
              "         D.通用名称 as 名称,C.规格,B.批号,to_char(B.效期,'yyyy-MM-dd') as 效期" & gstrSQL & ",A.发票金额 as 金额 " & _
              "    from 药品应付记录 A,药品收发记录 B,药品目录 C,药品信息 D,药品付款记录 E  " & _
              "    Where A.收发ID = B.ID And B.药品ID = C.药品ID And C.药名ID = D.药名ID  And A.付款序号=E.付款序号(+) and E.序号(+)=1 and E.审核日期 is null " & _
              "        and B.审核日期>=to_date('" & strBegin & "','yyyyMMdd') and B.审核日期<to_date('" & strEnd & "','yyyyMMdd') and A.供药单位ID=" & lngID & _
              "  order by B.NO,B.序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        mshDetail.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With mshDetail
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("入库单据号")), " ", rsTemp("入库单据号"))
            .TextMatrix(lngRow, 2) = IIf(IsNull(rsTemp("日期")), " ", rsTemp("日期"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("规格")), "", rsTemp("规格"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("单位")), "", rsTemp("单位"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rsTemp("批号")), "", rsTemp("批号"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rsTemp("效期")), "", rsTemp("效期"))
            .TextMatrix(lngRow, 8) = Format(rsTemp("数量"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(rsTemp("金额"), "###########0.00;-###########0.00; ; ")
                
            dblSum = dblSum + rsTemp("金额")
            lngRow = lngRow + 1
            
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, 0) = "合计"
            .TextMatrix(lngRow, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(dblSum, "###########0.00;-###########0.00; ; ")
        End If
        '再算单据的合计金额
        lngRow = 1
        Do While lngRow < .Rows - 1
            dblSum = Val(.TextMatrix(lngRow, 9))
            lngTemp = lngRow + 1
            
            Do While lngTemp < .Rows - 1
                If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) Then
                    dblSum = dblSum + Val(.TextMatrix(lngTemp, 9))
                Else
                    Exit Do
                End If
                lngTemp = lngTemp + 1
            Loop
            For lngCount = lngRow To lngTemp - 1
                .TextMatrix(lngCount, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            Next
            lngRow = lngTemp
        Loop
    End With
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
            .Row = 1
        End If
        
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
    
    If ActiveControl Is mshSum Then
        blnPrint = Not (mshSum.Rows = 2 And mshSum.TextMatrix(1, 0) = "")
    Else
        blnPrint = Not (mshDetail.Rows = 2 And mshDetail.TextMatrix(1, 0) = "")
    End If

    mnuFilePreView.Enabled = blnPrint
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
    Call zlWebForum(Me.hwnd)
End Sub

