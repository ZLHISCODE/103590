VERSION 5.00
Object = "*\A..\pzlTable.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Demo"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   11460
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdProperty 
      Caption         =   "属性对话框(&P)"
      Height          =   420
      Left            =   2385
      TabIndex        =   26
      Top             =   8460
      Width           =   1410
   End
   Begin VB.PictureBox picTMP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   9855
      ScaleHeight     =   690
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   945
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   9810
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.JPG|*.JPG|*.BMP|*.BMP|*.GIF|*.GIF|*.*|*.*"
   End
   Begin VB.Frame fraOptions 
      Height          =   8925
      Left            =   90
      TabIndex        =   20
      Top             =   0
      Width           =   2235
      Begin VB.PictureBox picBehaviourGroup 
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   60
         ScaleHeight     =   7815
         ScaleWidth      =   2115
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   765
         Width           =   2115
         Begin VB.CommandButton cmdFont 
            Caption         =   "改变单元格字体(&F)..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   0
            TabIndex        =   17
            Top             =   5265
            Width           =   2085
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "输出绘图结果(&X)..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   0
            TabIndex        =   19
            Top             =   6795
            Width           =   2085
         End
         Begin VB.CheckBox chkEditable 
            Appearance      =   0  'Flat
            Caption         =   "可编辑(&E)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   618
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkEnabled 
            Appearance      =   0  'Flat
            Caption         =   "可用(&N)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Top             =   315
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkSingleClickEdit 
            Appearance      =   0  'Flat
            Caption         =   "单击编辑(&S)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   921
            Width           =   1815
         End
         Begin VB.CheckBox chkHotTrack 
            Appearance      =   0  'Flat
            Caption         =   "热跟踪(&H)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   1224
            Width           =   1935
         End
         Begin VB.CheckBox chkBackground 
            Appearance      =   0  'Flat
            Caption         =   "背景图片(&B)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   1527
            Width           =   1815
         End
         Begin VB.CheckBox chkHighlightSelectedIcons 
            Appearance      =   0  'Flat
            Caption         =   "图标高亮(&H)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   1830
            Width           =   2055
         End
         Begin VB.CheckBox chkDrawFocusRect 
            Appearance      =   0  'Flat
            Caption         =   "焦点虚框(&F)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   2133
            Width           =   2055
         End
         Begin VB.CheckBox chkAlternateRowColour 
            Appearance      =   0  'Flat
            Caption         =   "间隔颜色(&T)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   2436
            Width           =   2055
         End
         Begin VB.CheckBox chkBlendSelection 
            Appearance      =   0  'Flat
            Caption         =   "选择半透明(&L)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   3345
            Width           =   2055
         End
         Begin VB.CheckBox chkCustomColours 
            Appearance      =   0  'Flat
            Caption         =   "单元格背景色(&C)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   2739
            Width           =   1815
         End
         Begin VB.CheckBox chkMergeCells 
            Appearance      =   0  'Flat
            Caption         =   "合并单元格(&M)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   3042
            Width           =   1815
         End
         Begin VB.CheckBox chkAutoHeight 
            Appearance      =   0  'Flat
            Caption         =   "自动行高(&A)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   3648
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkShowToolTips 
            Appearance      =   0  'Flat
            Caption         =   "显示提示文本(&I)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   3951
            Width           =   2055
         End
         Begin VB.CheckBox chkWordEllipsis 
            Appearance      =   0  'Flat
            Caption         =   "显示未完省略号(&Z)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   15
            Top             =   4557
            Width           =   2055
         End
         Begin VB.CheckBox chkSingleLine 
            Appearance      =   0  'Flat
            Caption         =   "单行文本(&U)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   4254
            Width           =   1815
         End
         Begin VB.CheckBox chkTabTrip 
            Appearance      =   0  'Flat
            Caption         =   "捕捉Tab键(&K)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   4860
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   1020
            Left            =   510
            TabIndex        =   18
            Top             =   5715
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1799
            ButtonWidth     =   609
            ButtonHeight    =   582
            Appearance      =   1
            Style           =   1
            ImageList       =   "imlAlign"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   9
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000010&
            Caption         =   " 属性"
            ForeColor       =   &H80000016&
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   60
         Picture         =   "frmTest.frx":0000
         ScaleHeight     =   540
         ScaleWidth      =   2115
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   180
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "注:本控件不支持滚动条"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   24
         Top             =   8640
         Width           =   1995
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10665
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4316
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":46B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":517E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5518
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":58B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6380
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":671A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":71E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7582
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":791C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8050
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":83EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8784
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9252
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":95EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9986
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":A0BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":A454
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":A7EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":AB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":AF22
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":B2BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":B656
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":B9F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":BD8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C124
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C182
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C23E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C29C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin zlTable.Table Table1 
      Height          =   5550
      Left            =   2385
      TabIndex        =   0
      Top             =   90
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   9790
      SingleLine      =   0   'False
   End
   Begin MSComctlLib.ImageList imlAlign 
      Left            =   2430
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C2FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C401
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C4CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C59B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C69D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C8A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":C966
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":CA64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBG 
      Height          =   3360
      Left            =   10890
      Picture         =   "frmTest.frx":CB2D
      Top             =   7560
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExport_Click()
    Dim strF As String
    picTMP.Width = Table1.Width
    picTMP.Height = Table1.Height
    Table1.DrawToDC picTMP.hDC
    picTMP.Picture = picTMP.Image
    dlgThis.FileName = ""
    dlgThis.ShowSave
    strF = dlgThis.FileName
    If strF <> "" Then
        SavePicture picTMP.Picture, strF
        MsgBox "表格输出为图片成功！" & vbCrLf & "文件名: " & strF
    End If
End Sub

Private Sub cmdProperty_Click()
    Table1.ShowProperty Me, Me.Table1, 3
End Sub

Private Sub Form_Load()
    Dim lKey As Long, i As Long, j As Long, Row As Long, Col As Long
    Row = 20
    Col = 6
    Randomize Timer

    With Table1
        .Init Row, Col
        .ImageList = Me.ImageList1
        .BorderColor = RGB(150, 150, 150)
        .BorderWidth = 2
        .GridLineColor = RGB(150, 150, 150)
        .GridLineWidth = 1
        .FontQuality = FQClearType
        .Redraw = False
        
        '用于显示调整列宽的参考线
        .hWndBound = Me.hWnd
        .OffsetX = .Left
        .OffsetY = .Top
        
        .ColWidth(1) = -600
        .ColWidth(2) = 0
        .ColWidth(3) = 2600
        .ColWidth(4) = -500
        .ColWidth(5) = 1000
        .ColWidth(6) = 3000
        
        .CellDetails 1, 1, "序号", , , , "序号", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        .CellDetails 1, 2, "标题", , , , "标题", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        .CellDetails 1, 3, "ID", , , , "ID", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        .CellDetails 1, 4, "图标", , , , "图标" & vbCrLf & "保护＋图标示例", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        .CellDetails 1, 5, "金额", , , , "金额" & vbCrLf & "这里用到了格式化FormatString属性", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        .CellDetails 1, 6, "说明", , , , "说明", , , True, HALignCentre, VALignVCentre, bFontBold:=True, oBackColor:=RGB(200, 200, 200)
        For i = 2 To 20
            .CellDetails i, 1, i - 1, "(0)", , , , , , True, HALignCentre, VALignVCentre
        Next i
        
        For i = 2 To 20
            For j = 1 To Col
                Select Case j
                Case 1
                    .CellDetails i, j, i - 1, "(0)", , , , , , True, HALignCentre, VALignVCentre
                Case 2
                    .CellDetails i, j, "中文 English"
                Case 3
                    .CellDetails i, j, "ID_FORMAT_FONT "
                Case 4
                    .CellDetails i, j, , , , Rnd() * 21 + 1, , , , True, HALignCentre, VALignVCentre
                Case 5
                    .CellDetails i, j, Format(Rnd() * 100, "0.00"), "￥#,0.00", , , , , , , HALignRight, VALignVCentre
                Case 6
                    .CellDetails i, j, "解释说明", , , , IIf(i = 2, "这里输入本单元格的说明文字" & vbCrLf & "支持换行符。", "")
                End Select
            Next
        Next i
        
        .Redraw = True
        .Refresh
    End With
End Sub

Private Sub chkAlternateRowColour_Click()
    If chkAlternateRowColour.Value = vbChecked Then
        Table1.AlternateRowBackColor = RGB(200, 255, 200)
    Else
        Table1.AlternateRowBackColor = -1
    End If
    Table1.Refresh False, False
End Sub

Private Sub chkAutoHeight_Click()
    Table1.AutoHeight = (chkAutoHeight.Value = vbChecked)
    Table1.Refresh False, Table1.AutoHeight
End Sub

Private Sub chkBackground_Click()
    If chkBackground.Value = vbChecked Then
        Table1.BackgroundPicture = imgBG.Picture
    Else
        Table1.BackgroundPicture = Nothing
    End If
End Sub

Private Sub chkBlendSelection_Click()
    If chkBlendSelection = vbChecked Then
        Table1.HighlightMode = HMFilledRectAlpha
    Else
        Table1.HighlightMode = HMFilledRectSolid
    End If
End Sub

Private Sub chkCustomColours_Click()
    If chkCustomColours.Value = vbChecked Then
        Table1.Cell(2, 6).BackColor = vbYellow
    Else
        Table1.Cell(2, 6).BackColor = -1
    End If
    Table1.Refresh False, False
End Sub

Private Sub chkDrawFocusRect_Click()
    Table1.DrawFocusRect = (chkDrawFocusRect.Value = vbChecked)
    Table1.Refresh False, False
End Sub

Private Sub chkEditable_Click()
    Table1.Editable = (chkEditable.Value = vbChecked)
    chkSingleClickEdit.Enabled = Table1.Editable
End Sub

Private Sub chkEnabled_Click()
    Table1.Enabled = (chkEnabled.Value = vbChecked)
End Sub

Private Sub chkHighlightSelectedIcons_Click()
    Table1.HighlightSelectedIcons = (chkHighlightSelectedIcons.Value = vbChecked)
    Table1.Refresh False, False
End Sub

Private Sub chkHotTrack_Click()
    Table1.HotTrack = (chkHotTrack.Value = vbChecked)
End Sub

Private Sub chkMergeCells_Click()
    If chkMergeCells.Value = vbChecked Then
        Table1.MergeSelectedCells
    Else
        Table1.DisMergeCells Table1.Row, Table1.Col
    End If
End Sub

Private Sub chkShowToolTips_Click()
    Table1.ShowToolTipText = (chkShowToolTips.Value = vbChecked)
End Sub

Private Sub chkSingleClickEdit_Click()
    Table1.SingleClickEdit = (chkSingleClickEdit.Value = vbChecked)
End Sub

Private Sub chkSingleLine_Click()
    Table1.SingleLine = (chkSingleLine.Value = vbChecked)
    chkWordEllipsis.Value = IIf(Table1.SingleLine = False, chkWordEllipsis.Value, vbUnchecked)
    chkWordEllipsis.Enabled = Table1.SingleLine
End Sub

Private Sub chkTabTrip_Click()
    Table1.TabKeyMoveNextCell = (chkTabTrip.Value = vbChecked)
End Sub

Private Sub chkWordEllipsis_Click()
    Table1.WordEllipsis = (chkWordEllipsis.Value = vbChecked)
End Sub

Private Sub cmdFont_Click()
    On Error GoTo LL
    Dim i As Long
    i = Table1.SelectedCellKey
    If i > 0 Then
        dlgThis.CancelError = True
        dlgThis.Flags = cdlCFBoth Or cdlCFEffects
        With Table1.Cells(i)
            dlgThis.FontBold = .FontBold
            dlgThis.FontItalic = .FontItalic
            dlgThis.FontName = .FontName
            dlgThis.FontSize = .FontSize
            dlgThis.FontStrikethru = .FontStrikeout
            dlgThis.FontUnderline = .FontUnderline
            dlgThis.Color = .ForeColor
            
            dlgThis.ShowFont
        
            .FontBold = dlgThis.FontBold
            .FontItalic = dlgThis.FontItalic
            .FontName = dlgThis.FontName
            .FontSize = dlgThis.FontSize
            .FontStrikeout = dlgThis.FontStrikethru
            .FontUnderline = dlgThis.FontUnderline
            .ForeColor = dlgThis.Color
        End With
        Table1.Refresh False, True, Table1.Cells(i).Row
    End If
LL:
End Sub

Private Sub Table1_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
'    chkMergeCells.Value = IIf(Table1.Cell(lRow, lCol).MergeInfo <> "", vbChecked, vbUnchecked)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    i = Table1.SelectedCellKey
    If i <= 0 Then Exit Sub
    Select Case Button.Index
    Case 1, 2, 3
        Table1.Cells(i).VAlignment = VALignTop
    Case 4, 5, 6
        Table1.Cells(i).VAlignment = VALignVCentre
    Case 7, 8, 9
        Table1.Cells(i).VAlignment = VALignBottom
    End Select
    Select Case Button.Index
    Case 1, 4, 7
        Table1.Cells(i).HAlignment = HALignLeft
    Case 2, 5, 8
        Table1.Cells(i).HAlignment = HALignCentre
    Case 3, 6, 9
        Table1.Cells(i).HAlignment = HALignRight
    End Select
    Table1.Refresh False, False, Table1.Cells(i).Row
End Sub










