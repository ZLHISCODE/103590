VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmPreView 
   BackColor       =   &H00808080&
   Caption         =   "打印预览"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmPreView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   11880
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   1245
      Left            =   5820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1500
      Width           =   285
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPreView.frx":014A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin MSComctlLib.ImageList Ils彩色 
      Left            =   585
      Top             =   1155
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
            Picture         =   "frmPreView.frx":09DE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":0BFA
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":0F16
            Key             =   "Margin"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":1612
            Key             =   "Dual"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":1D0E
            Key             =   "First"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":1F2A
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":2146
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":2362
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":257E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ils单色 
      Left            =   0
      Top             =   1125
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
            Picture         =   "frmPreView.frx":2798
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":29B4
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":2CD0
            Key             =   "Margin"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":33CC
            Key             =   "Dual"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":3AC8
            Key             =   "First"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":3CE4
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":3F00
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":411C
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreView.frx":4338
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar Coo标准 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1376
      BandCount       =   1
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      _CBWidth        =   11880
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "Tlb标准"
      MinHeight1      =   720
      Width1          =   11820
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin MSComctlLib.Toolbar Tlb标准 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   1270
         ButtonWidth     =   953
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "Ils单色"
         HotImageList    =   "Ils彩色"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Excel"
               Key             =   "Excel"
               Object.ToolTipText     =   "输出到Excel"
               Object.Tag             =   "Excel"
               ImageKey        =   "Excel"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "页边距"
               Key             =   "Margin"
               Object.ToolTipText     =   "页边距"
               Object.Tag             =   "页边距"
               ImageKey        =   "Margin"
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "双页"
               Key             =   "Dual"
               Description     =   "两页显示"
               Object.ToolTipText     =   "双页显示"
               Object.Tag             =   "Dual"
               ImageKey        =   "Dual"
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line1"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "首页"
               Key             =   "First"
               Description     =   "首条记录"
               Object.ToolTipText     =   "首页"
               Object.Tag             =   "首页"
               ImageKey        =   "First"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上页"
               Key             =   "Previous"
               Description     =   "上一条"
               Object.ToolTipText     =   "上页"
               Object.Tag             =   "上页"
               ImageKey        =   "Previous"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下页"
               Key             =   "Next"
               Description     =   "下一条"
               Object.ToolTipText     =   "下页"
               Object.Tag             =   "下页"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "末页"
               Key             =   "Last"
               Description     =   "末条"
               Object.ToolTipText     =   "末页"
               Object.Tag             =   "末页"
               ImageKey        =   "Last"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.ComboBox cmbPageNumber 
            Height          =   300
            Left            =   8400
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   210
            Width           =   1785
         End
         Begin VB.ComboBox cmbScale 
            Height          =   300
            Left            =   6150
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   210
            Width           =   2115
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3750
      Width           =   1245
   End
   Begin VB.PictureBox picCorner 
      Height          =   495
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   945
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2970
      Width           =   1005
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2325
      Index           =   1
      Left            =   1950
      ScaleHeight     =   2325
      ScaleWidth      =   3795
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3795
      Begin VB.Line linLeft 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   3360
         X2              =   3360
         Y1              =   30
         Y2              =   1650
      End
      Begin VB.Line linRight 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   2940
         X2              =   2940
         Y1              =   300
         Y2              =   1920
      End
      Begin VB.Line linFooter 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   270
         X2              =   3240
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line linDown 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   180
         X2              =   3150
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linHeader 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   270
         X2              =   3240
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line linUp 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   330
         X2              =   3300
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   570
         MousePointer    =   7  'Size N S
         TabIndex        =   14
         Top             =   210
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblFooter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   420
         MousePointer    =   7  'Size N S
         TabIndex        =   10
         Top             =   330
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   150
         MousePointer    =   7  'Size N S
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   1710
         MousePointer    =   9  'Size W E
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   2640
         MousePointer    =   9  'Size W E
         TabIndex        =   12
         Top             =   210
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   810
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2265
      End
   End
   Begin VB.Image imgStore 
      Height          =   300
      Left            =   210
      Picture         =   "frmPreView.frx":4552
      Top             =   3930
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape shpBack 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   1485
      Index           =   1
      Left            =   6120
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "页面设置(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&EXCEL"
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuToolBar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDual 
         Caption         =   "双页显示(&D)"
      End
      Begin VB.Menu mnuViewMargin 
         Caption         =   "页边距(&M)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFisrt 
         Caption         =   "首页(&F)"
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "上页(&P)"
      End
      Begin VB.Menu mnuViewNext 
         Caption         =   "下页(&N)"
      End
      Begin VB.Menu mnuViewLast 
         Caption         =   "末页(&L)"
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmPreView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnX As Boolean     '当为真时，HScoll滚动条的变化会放大一百倍
Dim mblnY As Boolean     '当为真时，VScoll滚动条的变化会放大一百倍
Dim msngX As Single      '鼠标按下时的X坐标
Dim msngY As Single      '鼠标按下时的Y坐标
Dim mlngStartPage As Long '本文件起始页
Dim mlngMaxPage As Long   '总页数
'本窗体是预览时的主窗口
'
'ShowPage               显示指定页
'

Public Event AfterPrint()

Private Sub cmbPageNumber_Click()
    Dim i As Integer
    Me.cmbScale.Enabled = False
    Me.cmbPageNumber.Enabled = False
    If gintShow = 1 Then
        If gintPage <> Val(Mid(cmbPageNumber.Text, 2)) Then
            gintPage = Val(Mid(cmbPageNumber.Text, 2))
            stbThis.Panels(2).Text = "共" & CStr(mlngMaxPage - mlngStartPage + 1) & "页，当前是第" & CStr(gintPage) & "页"
            Me.picDraw(1).Cls
            Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
            PrintPage gintPage
        End If
    Else
        If Val(Mid(cmbPageNumber.Text, 2)) > mlngMaxPage - mlngStartPage Then Exit Sub
        If Val(Mid(cmbPageNumber.Text, 2)) = mlngMaxPage - mlngStartPage Then
            cmbPageNumber.Text = "第" & CStr(mlngMaxPage - mlngStartPage + 1) & "页"
            Exit Sub
        End If
        If gintPage <> Val(Mid(cmbPageNumber.Text, 2)) Then
            gintPage = Val(Mid(cmbPageNumber.Text, 2))
            stbThis.Panels(2).Text = "共" & CStr(mlngMaxPage - mlngStartPage + 1) & "页，当前是第" & CStr(gintPage) & "页"
            Me.picDraw(1).Cls
            Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
            PrintPage gintPage
            Set gobjOutTo = Me.picDraw(2)
            Me.picDraw(2).Cls
            gintPage = gintPage + 1
            Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
            PrintPage gintPage
            Set gobjOutTo = Me.picDraw(1)
            gintPage = gintPage - 1
        End If
    End If
    Me.cmbScale.Enabled = True
    Me.cmbPageNumber.Enabled = True
End Sub

Private Sub cmbScale_Click()
    Select Case cmbScale.Text
        Case "原始大小"
           If gsngScale = 1 Then Exit Sub
           gsngScale = 1
        Case "最适合大小"
        
        Case Else
            gsngScale = Val(cmbScale.Text) / 100
    End Select
    cmbScale.Refresh
    Call ShowPage
    
End Sub
Private Sub ShowPage()
    '------------------------------------------------
    '功能： 显示指定页
    '参数：无
    '返回：无
    '------------------------------------------------

    picDraw(1).Cls
    picDraw(1).Width = gsngScaleWidth * gsngScale
    picDraw(1).Height = gsngScaleHeight * gsngScale
    picDraw(1).ScaleWidth = gsngScaleWidth
    picDraw(1).ScaleHeight = gsngScaleHeight
    shpBack(1).Width = picDraw(1).Width
    shpBack(1).Height = picDraw(1).Height
    stbThis.Panels(2).Text = "共" & CStr(mlngMaxPage - mlngStartPage + 1) & "页，当前是第" & CStr(gintPage) & "页"
    Form_Resize
    Me.cmbPageNumber.Enabled = False
    Me.cmbScale.Enabled = False
    
    Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
    PrintPage gintPage
    If gintShow = 2 Then
        picDraw(2).Cls
        Set gobjOutTo = picDraw(2)
        gintPage = gintPage + 1
        Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
        PrintPage gintPage
        gintPage = gintPage - 1
        Set gobjOutTo = picDraw(1)
    End If
    Me.cmbPageNumber.Enabled = True
    Me.cmbScale.Enabled = True
End Sub

Private Sub Form_Load()
    If gstrGrant <> "" Then
        stbThis.Panels(1).Picture = imgStore.Picture
        stbThis.Panels(1).Text = ""
    Else
        Call ApplyOEM(stbThis)
    End If
    
    '使窗口最大
    Set gobjOutTo = Me.picDraw(1)
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    mlngMaxPage = frmTendFileReader.GetPages
    mlngStartPage = frmTendFileReader.GetStartPage
    If mlngMaxPage - mlngStartPage + 1 = 1 Then
        mnuViewDual.Enabled = False
        Me.Tlb标准.Buttons("Dual").Enabled = False
    End If
    
    picDraw(1).Width = gsngScaleWidth
    picDraw(1).Height = gsngScaleHeight
    picDraw(1).ScaleWidth = gsngScaleWidth
    picDraw(1).ScaleHeight = gsngScaleHeight
    
    Dim blnTemp As Boolean
    blnTemp = HaveExcel
    mnuFileExcel.Enabled = blnTemp
    Tlb标准.Buttons("Excel").Enabled = blnTemp
    
    gsngScale = 1
    cmbScale.AddItem "原始大小"
    cmbScale.AddItem "250%"
    cmbScale.AddItem "200%"
    cmbScale.AddItem "150%"
    cmbScale.AddItem "75%"
    cmbScale.AddItem "50%"
    cmbScale.AddItem "25%"
    
    cmbScale.Text = "原始大小"
    
    Dim i As Integer
    For i = 1 To mlngMaxPage - mlngStartPage + 1
        cmbPageNumber.AddItem "第" & CStr(i) & "页"
    Next
    cmbPageNumber.ListIndex = 0
    
    '根据权限设置控件状态
    If InStr(1, ";" & gstrPrivs & ";", ";Excel输出;") = 0 Then
        mnuFileExcel.Visible = False
        Tlb标准.Buttons("Excel").Visible = False
    End If
    
    If InStr(1, ";" & gstrPrivs & ";", ";打印;") = 0 Then
        mnuFilePrint.Visible = False
        Tlb标准.Buttons("Print").Visible = False
    End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Dim sngTop As Single
   Dim sngBottom As Single
   
   If stbThis.Visible Then
       sngBottom = Me.ScaleHeight - stbThis.Height
   Else
       sngBottom = Me.ScaleHeight
   End If
   If Coo标准.Visible Then
       sngTop = Coo标准.Top + Coo标准.Height
   Else
       sngTop = Me.ScaleTop
   End If
    
   If gintShow = 1 Then
        If picDraw(1).Width + 400 > Me.ScaleWidth Then
            HScroll1.Visible = True
        Else
            HScroll1.Visible = False
        End If
    
   Else
        picDraw(2).Width = picDraw(1).Width
        picDraw(2).ScaleWidth = picDraw(1).ScaleWidth
        picDraw(2).Height = picDraw(1).Height
        picDraw(2).ScaleHeight = picDraw(1).ScaleHeight
        shpBack(2).Width = shpBack(1).Width
        shpBack(2).Height = shpBack(1).Height
        If picDraw(1).Width * 2 + 600 > Me.ScaleWidth Then
            HScroll1.Visible = True
        Else
            HScroll1.Visible = False
        End If
    
   End If
    If picDraw(1).Height + 400 > sngBottom - sngTop Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
     
    picCorner.Visible = HScroll1.Visible And VScroll1.Visible
    If picCorner.Visible Then
        picCorner.Height = HScroll1.Height
        picCorner.Width = VScroll1.Width
        picCorner.Left = Me.ScaleWidth - VScroll1.Width
        picCorner.Top = sngBottom - HScroll1.Height
    End If
    
    If HScroll1.Visible Then
         HScroll1.Left = Me.ScaleLeft
         HScroll1.Top = sngBottom - HScroll1.Height
         HScroll1.Width = IIf(picCorner.Visible, picCorner.Left - Me.ScaleLeft, Me.ScaleWidth - Me.ScaleLeft)
         If gintShow = 1 Then
            If Abs(Me.ScaleWidth - picDraw(1).Width - 500) < 30000 Then
               mblnX = False
               HScroll1.min = 200
               HScroll1.Max = Me.ScaleWidth - picDraw(1).Width - 500
            Else
               mblnX = True
               HScroll1.min = 2
               HScroll1.Max = (Me.ScaleWidth - picDraw(1).Width - 500) / 100
            End If
        Else
            If Abs(Me.ScaleWidth - picDraw(1).Width * 2 - 700) < 30000 Then
               mblnX = False
               HScroll1.min = 200
               HScroll1.Max = Me.ScaleWidth - picDraw(1).Width * 2 - 700
            Else
               mblnX = True
               HScroll1.min = 2
               HScroll1.Max = (Me.ScaleWidth - picDraw(1).Width * 2 - 700) / 100
            End If
        End If
        HScroll1.Value = HScroll1.min
        HScroll1.SmallChange = Abs(HScroll1.Max - HScroll1.min) / 10
        HScroll1.LargeChange = Abs(HScroll1.Max - HScroll1.min) / 2
        picDraw(1).Left = 200
        If gintShow = 2 Then picDraw(2).Left = picDraw(1).Left + picDraw(1).Width + 200
    Else
        If gintShow = 1 Then
            picDraw(1).Left = (Me.ScaleWidth - picDraw(1).Width + 60) / 2
        Else
            picDraw(1).Left = (Me.ScaleWidth - picDraw(1).Width * 2 + 260) / 2
            picDraw(2).Left = picDraw(1).Left + picDraw(1).Width + 200
        End If
    End If
    If VScroll1.Visible Then
         VScroll1.Left = Me.ScaleWidth - VScroll1.Width
         VScroll1.Top = sngTop
         VScroll1.Height = IIf(picCorner.Visible, picCorner.Top - VScroll1.Top, sngBottom - VScroll1.Top)
         
         If Abs(sngBottom - sngTop - picDraw(1).Height - 200) < 30000 Then
            mblnY = False
            VScroll1.min = sngTop + 200
            VScroll1.Max = sngBottom - sngTop - picDraw(1).Height - 200
         Else
            mblnY = True
            VScroll1.min = (sngTop + 200) / 100
            VScroll1.Max = (sngBottom - sngTop - picDraw(1).Height - 200) / 100
         End If
        VScroll1.Value = VScroll1.min
         
         VScroll1.SmallChange = Abs(VScroll1.Max - VScroll1.min) / 10
         VScroll1.LargeChange = Abs(VScroll1.Max - VScroll1.min) / 2
         picDraw(1).Top = sngTop + 200
         If gintShow = 2 Then picDraw(2).Top = picDraw(1).Top
    Else
        picDraw(1).Top = (sngBottom - sngTop - picDraw(1).Height + 60) / 2 + sngTop
         If gintShow = 2 Then picDraw(2).Top = picDraw(1).Top
    End If
    shpBack(1).Width = picDraw(1).Width
    shpBack(1).Height = picDraw(1).Height
    shpBack(1).Left = picDraw(1).Left + 60
    shpBack(1).Top = picDraw(1).Top + 60
    If gintShow = 2 Then
        shpBack(2).Width = picDraw(2).Width
        shpBack(2).Height = picDraw(2).Height
        shpBack(2).Left = picDraw(2).Left + 60
        shpBack(2).Top = picDraw(2).Top + 60
    End If
   Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload gfrmTemp
    Set gfrmTemp = Nothing
End Sub

Private Sub HScroll1_Change()
    
    If mblnX Then
       picDraw(1).Left = HScroll1.Value * 100!
    Else
       picDraw(1).Left = HScroll1.Value
    End If
    shpBack(1).Left = picDraw(1).Left + 60
    If gintShow = 2 Then
        picDraw(2).Left = picDraw(1).Left + picDraw(1).Width + 200
        shpBack(2).Left = picDraw(2).Left + 60
    End If
End Sub

Private Sub lblHeader_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngY = Y
End Sub

Private Sub lblHeader_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngY As Single
    If Button = 1 Then
        sngY = lblHeader.Top + Y / gsngScale - msngY
        If sngY > 0 And sngY < linFooter.Y1 Then
            linHeader.Y1 = sngY
            linHeader.Y2 = sngY
        End If
    End If
    stbThis.Panels(2).Text = "页眉位置：" & Format((linHeader.Y1 / conRatemmToTwip), "###0.00")
End Sub

Private Sub lblHeader_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHeader.Top = linHeader.Y1 - 30
End Sub

Private Sub lblFooter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngY = Y
End Sub

Private Sub lblFooter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngY As Single
    If Button = 1 Then
        sngY = lblFooter.Top + Y / gsngScale - msngY
        If sngY > linHeader.Y1 And sngY < gsngScaleHeight Then
            linFooter.Y1 = sngY
            linFooter.Y2 = sngY
        End If
    End If
    stbThis.Panels(2).Text = "页脚位置：" & Format((gsngScaleHeight - linFooter.Y1) / conRatemmToTwip, "###0.00")
End Sub

Private Sub lblFooter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblFooter.Top = linFooter.Y1 - 30
End Sub

Private Sub lblUp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngY = Y
End Sub

Private Sub lblUp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngY As Single
    If Button = 1 Then
        sngY = lblUp.Top + Y / gsngScale - msngY
        If sngY > 0 And sngY < linDown.Y1 Then
            linUp.Y1 = sngY
            linUp.Y2 = sngY
        End If
    End If
    stbThis.Panels(2).Text = "页上边距：" & Format((linUp.Y1 / conRatemmToTwip), "###0.00")
End Sub

Private Sub lblUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblUp.Top = linUp.Y1 - 30
End Sub

Private Sub lblDown_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngY = Y
End Sub

Private Sub lblDown_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngY As Single
    If Button = 1 Then
        sngY = lblDown.Top + Y / gsngScale - msngY
        If sngY > linUp.Y1 And sngY < gsngScaleHeight Then
            linDown.Y1 = sngY
            linDown.Y2 = sngY
        End If
    End If
    stbThis.Panels(2).Text = "页下边距：" & Format((gsngScaleHeight - linDown.Y1) / conRatemmToTwip, "###0.00")
End Sub

Private Sub lblDown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblDown.Top = linDown.Y1 - 30
End Sub

Private Sub lblLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngX = x
End Sub

Private Sub lblLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngX As Single
    If Button = 1 Then
        sngX = lblLeft.Left + x / gsngScale - msngX
        If sngX > 0 And sngX < linRight.X1 Then
            linLeft.X1 = sngX
            linLeft.X2 = sngX
        End If
    End If
    stbThis.Panels(2).Text = "页左边距：" & Format((linLeft.X1 / conRatemmToTwip), "###0.00")
End Sub

Private Sub lblLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblLeft.Left = linLeft.X1 - 30
End Sub

Private Sub lblRight_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    msngX = x
End Sub

Private Sub lblRight_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sngX As Single
    If Button = 1 Then
        sngX = lblRight.Left + x / gsngScale - msngX
        If sngX > linLeft.X1 And sngX < gsngScaleWidth Then
            linRight.X1 = sngX
            linRight.X2 = sngX
        End If
    End If
    stbThis.Panels(2).Text = "页右边距：" & Format((gsngScaleWidth - linRight.X1) / conRatemmToTwip, "###0.00")
End Sub

Private Sub lblRight_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblRight.Left = linRight.X1 - 30
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuFileExcel_Click()
    Dim frmExcel As New frmOutExcel
    
    If gstrGrant <> "" Then
        MsgBox "试用或测试版本不能使用该功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not HaveExcel Then
        MsgBox "请安装好Excel或WPS(ET)后，再使用本功能。", vbCritical + vbOKOnly, gstrSysName
        Exit Sub
    End If
    gobjSend.Header = gstrHeader
    gobjSend.Footer = gstrFooter
    frmExcel.Show 1
    Set frmExcel = Nothing
End Sub

Private Sub mnuFilePrint_Click()
    Dim blnPrint As Boolean
    
    Dim frmPrintTemp As New frmPrint
    blnPrint = frmPrintTemp.PrintData
    Set frmPrintTemp = Nothing
    Set gobjOutTo = picDraw(1)
    
    RaiseEvent AfterPrint
End Sub

Private Sub mnuFileSet_Click()
'    If frmPageSet.ShowSet Then
'        'zlPutPrinterSet
'        ReDim gsngPrintedWidth(1 To gintTotalCol)
'        gsngHeight = gsngScaleHeight - (gsngUp + gsngDown) * conRatemmToTwip - gsngTitle - gsngDownAppRow - gsngUpAppRow - gsngFixRow - 2 * conLineHigh
'        gsngWidth = gsngScaleWidth - (gsngLeft + gsngRight) * conRatemmToTwip - gsngFixCol - 2 * conLineWide
'        Call CalculateRC
'        If gintColTotal * gintRowTotal = 1 Then
'            mnuViewDual.Enabled = False
'            Me.Tlb标准.Buttons("Dual").Enabled = False
'        Else
'            mnuViewDual.Enabled = True
'            Me.Tlb标准.Buttons("Dual").Enabled = True
'        End If
'
'        cmbPageNumber.Clear
'        Dim i As Integer
'        For i = 1 To gintColTotal * gintRowTotal
'            cmbPageNumber.AddItem "第" & CStr(i) & "页"
'        Next
'        If gintPage > gintColTotal * gintRowTotal Then gintPage = gintColTotal * gintRowTotal
'        Call ShowPage
'        cmbPageNumber.Text = "第" & CStr(gintPage) & "页"
'    End If
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuStatusBar_Click()
    mnuStatusBar.Checked = Not mnuStatusBar.Checked
    stbThis.Visible = mnuStatusBar.Checked
    Form_Resize
End Sub

Private Sub mnuViewDual_Click()
    Dim blnYes As Boolean '为真表示按下了
    
    If mlngMaxPage - mlngStartPage + 1 < 2 Then Exit Sub
    mnuViewDual.Checked = Not mnuViewDual.Checked
    blnYes = mnuViewDual.Checked
    If blnYes Then
        Tlb标准.Buttons("Dual").Value = tbrPressed
        gintShow = 2
        Load Me.picDraw(2)
        Load Me.shpBack(2)
        Me.picDraw(2).AutoRedraw = True
        Me.picDraw(2).Visible = True
        Me.shpBack(2).Visible = True
    Else
        Tlb标准.Buttons("Dual").Value = tbrUnpressed
        gintShow = 1
        Unload Me.picDraw(2)
        Unload Me.shpBack(2)
    End If
    If gintPage = mlngMaxPage - mlngStartPage + 1 Then
        gintPage = gintPage - 1
        cmbPageNumber.Text = "第" & CStr(gintPage) & "页"
    End If
    Form_Resize
    Me.cmbPageNumber.Enabled = False
    Me.cmbScale.Enabled = False

    stbThis.Panels(2).Text = "共" & CStr(mlngMaxPage - mlngStartPage + 1) & "页，当前是第" & CStr(gintPage) & "页"
    Me.picDraw(1).Cls
    Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
    PrintPage gintPage
    If blnYes Then
        Me.picDraw(2).Cls
        Set gobjOutTo = Me.picDraw(2)
        gintPage = gintPage + 1
        Call frmTendFileReader.ShowPage(gintPage + mlngStartPage - 1)
        PrintPage gintPage
        gintPage = gintPage - 1
        Set gobjOutTo = Me.picDraw(1)
    End If
    Me.cmbPageNumber.Enabled = True
    Me.cmbScale.Enabled = True
End Sub

Private Sub mnuViewFisrt_Click()
    Dim intPage As Integer
    If gintPage <> 1 Then
        intPage = 1
        cmbPageNumber.Text = "第" & CStr(intPage) & "页"
    End If
End Sub

Private Sub mnuViewFlash_Click()
    Me.picDraw(1).Refresh
End Sub

Private Sub mnuViewLast_Click()
    Dim intPage As Integer
    If gintShow = 1 Then
        If gintPage <> mlngMaxPage - mlngStartPage + 1 Then
            intPage = mlngMaxPage - mlngStartPage + 1
            cmbPageNumber.Text = "第" & CStr(intPage) & "页"
        End If
    Else
        If gintPage < mlngMaxPage - mlngStartPage + 1 Then
            intPage = mlngMaxPage - mlngStartPage
            cmbPageNumber.Text = "第" & CStr(intPage) & "页"
        End If
    End If
End Sub

Private Sub mnuViewMargin_Click()
    
'    Dim blnYes As Boolean '为真表示按下了
'    mnuViewMargin.Checked = Not mnuViewMargin.Checked
'    blnYes = mnuViewMargin.Checked
'    Tlb标准.Buttons("Margin").Value = IIf(blnYes, tbrPressed, tbrUnpressed)
'
'    linFooter.Visible = blnYes
'    linHeader.Visible = blnYes
'    linLeft.Visible = blnYes
'    linRight.Visible = blnYes
'    linUp.Visible = blnYes
'    linDown.Visible = blnYes
'
'    lblFooter.Visible = blnYes
'    lblHeader.Visible = blnYes
'    lblLeft.Visible = blnYes
'    lblRight.Visible = blnYes
'    lblUp.Visible = blnYes
'    lblDown.Visible = blnYes
'
'    blnYes = Not blnYes '为真表示可用
'
'    Tlb标准.Buttons("Print").Enabled = blnYes
'    Tlb标准.Buttons("Excel").Enabled = blnYes
'    Tlb标准.Buttons("Excel").Enabled = blnYes
'    Tlb标准.Buttons("Dual").Enabled = False
'    Tlb标准.Buttons("Next").Enabled = blnYes
'    Tlb标准.Buttons("Last").Enabled = blnYes
'    Tlb标准.Buttons("Previous").Enabled = blnYes
'
'    mnuFileExcel.Enabled = blnYes
'    mnuFilePrint.Enabled = blnYes
'    mnuFileSet.Enabled = blnYes
'    mnuViewDual.Enabled = False
'    mnuViewFisrt.Enabled = blnYes
'    mnuViewLast.Enabled = blnYes
'    mnuViewNext.Enabled = blnYes
'    mnuViewPrevious.Enabled = blnYes
'
'    cmbPageNumber.Enabled = blnYes
'    cmbScale.Enabled = blnYes
'
'    If Not blnYes Then
'        linFooter.X1 = 0: linFooter.X2 = picDraw(1).ScaleWidth
'        linFooter.Y1 = gsngScaleHeight - gsngFooter * conRatemmToTwip: linFooter.Y2 = gsngScaleHeight - gsngFooter * conRatemmToTwip
'
'        linHeader.X1 = 0: linHeader.X2 = picDraw(1).ScaleWidth
'        linHeader.Y1 = gsngHeader * conRatemmToTwip: linHeader.Y2 = gsngHeader * conRatemmToTwip
'
'        linUp.X1 = 0: linUp.X2 = picDraw(1).ScaleWidth
'        linUp.Y1 = gsngUp * conRatemmToTwip: linUp.Y2 = gsngUp * conRatemmToTwip
'
'        linDown.X1 = 0: linDown.X2 = picDraw(1).ScaleWidth
'        linDown.Y1 = gsngScaleHeight - gsngDown * conRatemmToTwip: linDown.Y2 = gsngScaleHeight - gsngDown * conRatemmToTwip
'
'        linLeft.Y1 = 0: linLeft.Y2 = picDraw(1).ScaleHeight
'        linLeft.X1 = gsngLeft * conRatemmToTwip: linLeft.X2 = gsngLeft * conRatemmToTwip
'
'        linRight.Y1 = 0: linRight.Y2 = picDraw(1).ScaleHeight
'        linRight.X1 = gsngScaleWidth - gsngRight * conRatemmToTwip: linRight.X2 = gsngScaleWidth - gsngRight * conRatemmToTwip
'
'        lblFooter.Left = 0: lblFooter.Width = picDraw(1).ScaleWidth
'        lblFooter.Top = linFooter.Y1 - 30
'
'        lblHeader.Left = 0: lblHeader.Width = picDraw(1).ScaleWidth
'        lblHeader.Top = linHeader.Y1 - 30
'
'        lblUp.Left = 0: lblUp.Width = picDraw(1).ScaleWidth
'        lblUp.Top = linUp.Y1 - 30
'
'        lblDown.Left = 0: lblDown.Width = picDraw(1).ScaleWidth
'        lblDown.Top = linDown.Y1 - 30
'
'        lblLeft.Top = 0: lblLeft.Height = picDraw(1).ScaleHeight
'        lblLeft.Left = linLeft.X1 - 30
'
'        lblRight.Top = 0: lblRight.Height = picDraw(1).ScaleHeight
'        lblRight.Left = linRight.X1 - 30
'    Else
'        If MsgBox("保存刚才的设置吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'            gsngHeader = linHeader.Y1 / conRatemmToTwip
'            gsngFooter = (gsngScaleHeight - linFooter.Y1) / conRatemmToTwip
'            gsngUp = linUp.Y1 / conRatemmToTwip
'            gsngDown = (gsngScaleHeight - linDown.Y1) / conRatemmToTwip
'            gsngLeft = linLeft.X1 / conRatemmToTwip
'            gsngRight = (gsngScaleWidth - linRight.X1) / conRatemmToTwip
'
'            gobjSend.EmptyDown = gsngDown
'            gobjSend.EmptyLeft = gsngLeft
'            gobjSend.EmptyRight = gsngRight
'            gobjSend.EmptyUp = gsngUp
'
'            ReDim gsngPrintedWidth(1 To gintTotalCol)
'            gsngHeight = gsngScaleHeight - (gsngUp + gsngDown) * conRatemmToTwip - gsngTitle - gsngDownAppRow - gsngUpAppRow - gsngFixRow - 2 * conLineHigh
'            gsngWidth = gsngScaleWidth - (gsngLeft + gsngRight) * conRatemmToTwip - gsngFixCol - 2 * conLineWide
'            Call CalculateRC
'
'            cmbPageNumber.Clear
'            Dim i As Integer
'            For i = 1 To gintColTotal * gintRowTotal
'                cmbPageNumber.AddItem "第" & CStr(i) & "页"
'            Next
'            If gintPage > gintColTotal * gintRowTotal Then gintPage = gintColTotal * gintRowTotal
'            Call ShowPage
'            cmbPageNumber.Text = "第" & CStr(gintPage) & "页"
'        End If
'
'        Dim blnTemp As Boolean
'        blnTemp = HaveExcel
'        mnuFileExcel.Enabled = blnTemp
'        Tlb标准.Buttons("Excel").Enabled = blnTemp
'
'        If gintColTotal * gintRowTotal = 1 Then
'            mnuViewDual.Enabled = False
'            Me.Tlb标准.Buttons("Dual").Enabled = False
'        Else
'            mnuViewDual.Enabled = True
'            Me.Tlb标准.Buttons("Dual").Enabled = True
'        End If
'    End If
'    stbThis.Panels(2).Text = "共" & CStr(gintColTotal * gintRowTotal) & "页，当前是第" & CStr(gintPage) & "页"
End Sub

Private Sub mnuViewNext_Click()
    Dim intPage As Integer
    If gintShow = 1 Then
        If gintPage < mlngMaxPage - mlngStartPage + 1 Then
            intPage = gintPage + 1
            cmbPageNumber.Text = "第" & CStr(intPage) & "页"
        End If
    Else
        If gintPage < mlngMaxPage - mlngStartPage Then
            intPage = gintPage + 1
            cmbPageNumber.Text = "第" & CStr(intPage) & "页"
        End If
    End If
End Sub

Private Sub mnuViewPrevious_Click()
    Dim intPage As Integer
    If gintPage > 1 Then
        intPage = gintPage - 1
        cmbPageNumber.Text = "第" & CStr(intPage) & "页"
    End If
End Sub

Private Sub mnuViewStand_Click()
    mnuViewStand.Checked = Not mnuViewStand.Checked
    Coo标准.Visible = mnuViewStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewText_Click()
    mnuViewText.Checked = Not mnuViewText.Checked
    Dim btnTemp As Object
    For Each btnTemp In Tlb标准.Buttons
        If mnuViewText.Checked Then
            btnTemp.Caption = btnTemp.Tag
        Else
            btnTemp.Caption = ""
        End If
    Next
    Coo标准.Bands(1).MinHeight = Tlb标准.Height
    cmbPageNumber.Top = (cmbPageNumber.Container.Height - cmbPageNumber.Height) / 2
    cmbScale.Top = cmbPageNumber.Top
    Form_Resize
End Sub

Private Sub Tlb标准_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuQuit_Click
        Case "Excel"
            mnuFileExcel_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Margin"
            mnuViewMargin_Click
        Case "Dual"
            mnuViewDual_Click
        Case "First"
            mnuViewFisrt_Click
        Case "Previous"
            mnuViewPrevious_Click
        Case "Next"
            mnuViewNext_Click
        Case "Last"
            mnuViewLast_Click
    End Select
End Sub

Private Sub VScroll1_Change()
    
    If Not mblnY Then
        picDraw(1).Top = VScroll1.Value
    Else
        picDraw(1).Top = 100! * VScroll1.Value
    End If
    shpBack(1).Top = picDraw(1).Top + 60
    If gintShow = 2 Then
        picDraw(2).Top = picDraw(1).Top
        shpBack(2).Top = picDraw(2).Top + 60
    End If
End Sub
