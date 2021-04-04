VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPatiFileQry 
   Caption         =   "病历检索"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   Icon            =   "frmPatiFileQry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox imgY 
      BackColor       =   &H00808080&
      Height          =   3375
      Index           =   0
      Left            =   4080
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.ImageList iLsTree32 
      Left            =   840
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":058A
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":0E64
            Key             =   "Attr"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMain 
      Left            =   1200
      Top             =   5800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":117E
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":139A
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":15B6
            Key             =   "新增"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":17D2
            Key             =   "插入"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":19EE
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":1C0A
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":1E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2042
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":225E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2478
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":28B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":2F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3604
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":3F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4132
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":48AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4CE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   3000
      Top             =   5920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":4EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":511A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":533A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":555A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":577A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":599A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":5FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6214
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6434
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6654
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6874
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":6CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":73A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":7ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8650
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":886A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8A84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   10740
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinHeight1      =   660
      Width1          =   9000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   660
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilstbrMain"
         HotImageList    =   "ilstbrMainHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "打印预览医嘱本"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印医嘱本"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "删除"
               Key             =   "删除病历"
               Description     =   "病历"
               Object.ToolTipText     =   "删除当前病历"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "归档"
               Key             =   "归档"
               Description     =   "病历"
               Object.ToolTipText     =   "将病历归档保存"
               Object.Tag             =   "归档"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找当前科的病人"
               Object.Tag             =   "查找"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_4"
               Description     =   "不用"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "查看"
               Key             =   "查看"
               Description     =   "不用"
               Object.ToolTipText     =   "变换显示图标方式"
               Object.Tag             =   "查看"
               ImageIndex      =   12
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出医嘱管理"
               Object.Tag             =   "退出"
               ImageIndex      =   14
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   0
      Top             =   5000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":8C9E
            Key             =   "门诊"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":9238
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":97D2
            Key             =   "护理"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":9D6C
            Key             =   "文书"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFileQry.frx":A306
            Key             =   "单据"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   200
      Left            =   2280
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7365
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiFileQry.frx":A8A0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picFile 
      Height          =   3735
      Left            =   4800
      ScaleHeight     =   3675
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
      Begin zl9CISCore.ctrlPatientFile ProFile1 
         Height          =   5175
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9128
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFind 
         Caption         =   "查找病历(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "病历预览(&L)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "病历打印(&Y)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuParamSet 
         Caption         =   "参数设置(&M)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuPatiRec 
      Caption         =   "病历(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuOrder_Edit 
         Caption         =   "修改病历(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "删除病历(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_File 
         Caption         =   "病历归档(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Undo 
         Caption         =   "病历作废(&U)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Print 
         Caption         =   "病历打印(&P)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuToolbar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu v1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIconOrder 
         Caption         =   "查看方式(&I)"
         Visible         =   0   'False
         Begin VB.Menu mnuIcon 
            Caption         =   "大图标(&G)"
            Index           =   0
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "小图标(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "列表(&L)"
            Index           =   2
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "详细资料(&D)"
            Checked         =   -1  'True
            Index           =   3
         End
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "病人信息(&I)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewHist 
         Caption         =   "病史分析(&A)"
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmPatiFileQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strPrivs As String       '用户具有本程序的具体权限

'查询条件串：病历种类||病人姓名||性别||最小年龄||最大年龄||医生||日期下限||日期上限||病历内容
Private strQuery As String
Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(frmParent As Object, Optional ByVal ModalWindow As Boolean = True)
    On Error Resume Next
    Set objParentForm = frmParent
    Me.Show IIf(ModalWindow, 1, 0), frmParent
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then Exit Sub
    
    Me.Tag = ""
    If Len(strQuery) > 0 Then
        ListItem
    End If
    
    If lvwItem.ListItems.Count > 0 Then lvwItem.ListItems(1).Selected = True: lvwItem_ItemClick lvwItem.SelectedItem
End Sub

Private Sub mnuExcel_Click()
    zlRptPrint 3
End Sub

Private Sub mnuFile_Preview_Click()
    Dim frmPreview As frmCasePrint
    Dim FileID As Long, PatientID As String, CheckID As Variant, FileType As Integer
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    FileID = Mid(Me.lvwItem.SelectedItem.Key, 4)
    On Error Resume Next
    zlDatabase.OpenRecordset rsTmp, "Select 病人ID,主页ID,挂号单,病历种类 From 病人病历记录 Where ID=" & FileID, Me.Caption
    If rsTmp.EOF Then Exit Sub
    PatientID = rsTmp(0): FileType = rsTmp(3)
        
    Set frmPreview = New frmCasePrint
    PrintOutCase Me, frmPreview, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
    frmPreview.Preview Me, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
End Sub

Private Sub mnuFile_Print_Click()
    Dim frmPreview As frmCasePrint
    Dim FileID As Long, PatientID As String, CheckID As Variant, FileType As Integer
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    FileID = Mid(Me.lvwItem.SelectedItem.Key, 4)
    On Error Resume Next
    zlDatabase.OpenRecordset rsTmp, "Select 病人ID,主页ID,挂号单,病历种类 From 病人病历记录 Where ID=" & FileID, Me.Caption
    If rsTmp.EOF Then Exit Sub
    PatientID = rsTmp(0): FileType = rsTmp(3)
        
    intPage = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(Me) Then
        MsgBox "打印机初始化失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    PrintOutCase Me, Printer, FileType, True, -1 * FileID, CLng(PatientID), "", False, 0, 1
    'WinNT自定义纸张处理
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(Me)
End Sub

Private Sub mnuFind_Click()
    Dim strTmp As String
    
    strTmp = strQuery
    frmPatiFileQry1.GetQueryString Me, strTmp
    If Len(strTmp) > 0 Then
        strQuery = strTmp
        
        ListItem
        If lvwItem.ListItems.Count > 0 Then lvwItem.ListItems(1).Selected = True
        lvwItem_ItemClick lvwItem.SelectedItem
    End If
End Sub

Private Sub mnuPreview_Click()
    zlRptPrint 2
End Sub

Private Sub mnuPrint_Click()
    zlRptPrint 1
End Sub

Private Sub mnuPrintSet_Click()
    frmPrintSet.Show vbModal, Me
End Sub

Private Sub mnuRefresh_Click()
    On Error Resume Next
    
    ListItem
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picFile_Resize()
    On Error Resume Next
    With ProFile1
        .Left = 0: .Top = 0
        .Width = picFile.ScaleWidth
        .Height = picFile.ScaleHeight
        
        If .Width > picFile.ScaleWidth Then Me.Width = .Width
        If .Height > picFile.ScaleHeight Then Me.Height = .Height + picFile.Top
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "预览"
            mnuPreview_Click
        Case "打印"
            mnuPrint_Click
        Case "查找"
            mnuFind_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
            mnuExit_Click
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_病历", "病历", 2000
        .Add , "_ID", "ID", 0
        .Add , "_日期", "日期", 1200
        .Add , "_姓名", "姓名", 800
        .Add , "_性别", "性别", 500
    End With
    With Me.lvwItem
        .ColumnHeaders("_病历").Position = 3
        .SortKey = .ColumnHeaders("_日期").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    
    '---------权限控制-------------
    strPrivs = gstrPrivs
    
    '读取保存的查询条件设置
    strQuery = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查询条件", "")
    ShowQryString strQuery
    
    Me.Tag = "Loading"
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.cbrMain.Visible, Me.cbrMain.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    imgY(0).Top = lngTools
    imgY(0).Height = Me.ScaleHeight - lngStatus - imgY(0).Top
    
    With lvwItem
        .Left = 0
        .Top = imgY(0).Top
        .Width = imgY(0).Left
        .Height = imgY(0).Height
    End With
    With picFile
        .Left = imgY(0).Left + imgY(0).Width: .Top = imgY(0).Top
        .Height = Me.ScaleHeight - lngStatus - .Top: .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存查询条件设置
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查询条件", strQuery
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgY_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    imgY(Index).Left = imgY(Index).Left + x
End Sub

Private Sub imgY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    Select Case Index
        Case 0
            If imgY(0).Left < 2000 Then imgY(0).Left = 2000
            If Me.ScaleWidth - imgY(0).Left < 4000 Then imgY(0).Left = Me.ScaleWidth - 4000
    End Select

    Form_Resize
End Sub

Private Sub lvwItem_DblClick()
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowMenu
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "显示病历："
    If Item Is Nothing Then
        ProFile1.ShowFile "", , , , -1, , , Me.prbRefresh '清除病历内容
    Else
        ProFile1.ShowFile Mid(Item.Key, 4), , , , , , , Me.prbRefresh
    End If
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    ShowQryString strQuery
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .SortKey = ColumnHeader.Index - 1: .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    End With
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuIcon_Click(Index As Integer)
'    lvwItem.View = Index
'    SetViewCheck lvwItem.View
End Sub

Private Sub mnuStatus_Click()
    Me.mnuStatus.Checked = Not Me.mnuStatus.Checked
    Me.stbThis.Visible = Me.mnuStatus.Checked
    Form_Resize
End Sub

Private Sub mnuToolbarStand_Click()
    Me.mnuToolbarStand.Checked = Not Me.mnuToolbarStand.Checked
    Me.cbrMain.Visible = Me.mnuToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuToolbarText_Click()
    Dim i As Integer
    Me.mnuToolbarText.Checked = Not Me.mnuToolbarText.Checked
    If Me.mnuToolbarText.Checked Then
        For i = 1 To Me.tbrMain.Buttons.Count
            Me.tbrMain.Buttons(i).Caption = Me.tbrMain.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tbrMain.Buttons.Count
            Me.tbrMain.Buttons(i).Caption = ""
        Next
    End If
    Me.cbrMain.Bands(1).MINHEIGHT = Me.tbrMain.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewHist_Click()
    On Error Resume Next
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    Call frmRcdAnalyse.ShowMe(1, Me, CLng(lvwItem.SelectedItem.SubItems(1)))
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "大图标(&G)"
        mnuIcon_Click 0
    Case "小图标(&M)"
        mnuIcon_Click 1
    Case "列表(&L)"
        mnuIcon_Click 2
    Case "详细资料(&D)"
        mnuIcon_Click 3
    End Select
End Sub

Private Sub ListItem()
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim iNum As Long
    Dim strWhereClause As String
    Dim aQryString() As String
    
    lvwItem.ListItems.Clear
    
    strWhereClause = ""
    If Len(strQuery) > 0 Then
        aQryString = Split(strQuery, "||")
        
        If Len(aQryString(0)) > 0 And aQryString(0) <> "0" Then strWhereClause = strWhereClause + " And a.病历种类=" + aQryString(0)
        If Len(aQryString(1)) > 0 Then strWhereClause = strWhereClause + " And b.姓名 Like '%" + Replace(aQryString(1), "'", "''") + "%'"
        If Len(aQryString(2)) > 0 And aQryString(2) <> "0" Then strWhereClause = strWhereClause + " And b.性别='" + aQryString(2) + "'"
        If Len(aQryString(3)) > 0 Then strWhereClause = strWhereClause + " And b.年龄>=" + aQryString(3)
        If Len(aQryString(4)) > 0 Then strWhereClause = strWhereClause + " And b.年龄<=" + aQryString(4)
        If Len(aQryString(5)) > 0 Then strWhereClause = strWhereClause + " And a.书写人 Like '%" + Replace(aQryString(5), "'", "''") + "%'"
        If Len(aQryString(6)) > 0 Then strWhereClause = strWhereClause + " And a.书写日期>=To_Date('" + aQryString(6) + "','yyyy-mm-dd')"
        If Len(aQryString(7)) > 0 Then strWhereClause = strWhereClause + " And a.书写日期<=To_Date('" + aQryString(7) + " 23:59:59','yyyy-mm-dd hh24:mi:ss')"
'        If Len(aQryString(8)) > 0 Then strWhereClause = strWhereClause + " And (d.内容 Like '%" + aQryString(8) + "%' or " + _
'            "e.所见内容 Like '%" + aQryString(8) + "%')"
            
        If Len(strWhereClause) > 0 Then strWhereClause = Mid(strWhereClause, 6)
    End If
    
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在检索："
    If Len(aQryString(8)) = 0 Then
        On Error GoTo QryError
        zlDatabase.OpenRecordset rsTmp, "Select a.ID,a.病人ID,To_Char(a.书写日期,'yyyy-mm-dd'),b.姓名,nvl(b.性别,' ')," + _
            "decode(a.病历种类,1,'门诊',2,'住院',3,'护理',4,'文书','单据'),a.病历名称 From 病人病历记录 a,病人信息 b Where " + _
            IIf(Len(strWhereClause) = 0, "", strWhereClause + " And ") + "a.病历种类>0 And a.病人ID=b.病人ID Order By a.书写日期,a.病人ID", Me.Caption
    Else
        On Error GoTo QryError
        zlDatabase.OpenRecordset rsTmp, "Select Distinct a.ID,a.病人ID,To_Char(a.书写日期,'yyyy-mm-dd'),b.姓名,nvl(b.性别,' ')," + _
            "decode(a.病历种类,1,'门诊',2,'住院',3,'护理',4,'文书','单据'),a.病历名称 From 病人病历记录 a,病人信息 b,病人病历内容 c Where " + _
            IIf(Len(strWhereClause) = 0, "", strWhereClause + " And ") + "a.病历种类>0 And a.病人ID=b.病人ID And a.ID=c.病历记录ID And c.ID In " + _
            "(Select 病历ID From 病人病历文本段 Where 内容 Like '%" + Replace(aQryString(8), "'", "''") + _
            "%' Union Select 病历ID From 病人病历所见单 Where 控件类 In (2,Null) And 所见内容 Like '%" + Replace(aQryString(8), "'", "''") + _
            "%') Order By To_Char(a.书写日期,'yyyy-mm-dd'),b.姓名", Me.Caption
    End If
    
    prbRefresh.Value = 50
    iNum = 0
    Do While Not rsTmp.EOF
        Set tmpItem = lvwItem.ListItems.Add(, "Key" & rsTmp(0), rsTmp(6))
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_ID").Index - 1) = rsTmp(1)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_日期").Index - 1) = rsTmp(2)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_姓名").Index - 1) = rsTmp(3)
        tmpItem.SubItems(Me.lvwItem.ColumnHeaders("_性别").Index - 1) = rsTmp(4)
        tmpItem.Icon = CStr(rsTmp(5)): tmpItem.SmallIcon = CStr(rsTmp(5))
        
        iNum = iNum + 1
        prbRefresh.Value = 50 + CLng(50 * iNum / rsTmp.RecordCount)
        rsTmp.MoveNext
    Loop
    Me.stbThis.Panels(3).Text = "病历记录：" + IIf(iNum = 0, "无", iNum & "条")
    prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    ShowQryString strQuery
    
    ShowMenu
    Exit Sub
QryError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub

Private Sub ShowMenu()
'    mnuOrder_Jz.Enabled = False
'    mnuOrder_Qx.Enabled = False
'    mnuOrder_Wc.Enabled = False
'    mnuOrder_Hf.Enabled = False
'    Select Case iPatiType
'        Case 0
'            If Not lvwItem(0).SelectedItem Is Nothing Then mnuOrder_Jz.Enabled = True
'        Case 1
'            If Not lvwItem(1).SelectedItem Is Nothing Then
'                mnuOrder_Qx.Enabled = True
'                mnuOrder_Wc.Enabled = True
'            End If
'        Case 2
'            If Not lvwItem(2).SelectedItem Is Nothing Then mnuOrder_Hf.Enabled = True
'    End Select
'
'    tbrMain.Buttons("接诊").Enabled = mnuOrder_Jz.Enabled
'    tbrMain.Buttons("完成").Enabled = mnuOrder_Wc.Enabled
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItem
    objPrint.Title.Text = "病历清单"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
'显示查询条件
Private Sub ShowQryString(ByVal strQry As String)
'查询条件串：病历种类||病人姓名||性别||最小年龄||最大年龄||医生||日期下限||日期上限||病历内容
    Dim aQryString() As String
    
    If Len(Trim(strQry)) = 0 Then Me.stbThis.Panels(2).Text = "查询条件  未设置": Exit Sub
    
    aQryString = Split(strQry, "||")
    With Me.stbThis.Panels(2)
        .Text = ""
        If Len(aQryString(0)) > 0 And aQryString(0) <> "0" Then
            Select Case aQryString(0)
                Case 1
                    .Text = .Text + "门诊病历，"
                Case 2
                    .Text = .Text + "住院病历，"
                Case 3
                    .Text = .Text + "护理记录，"
                Case 4
                    .Text = .Text + "诊断文书，"
                Case 5
                    .Text = .Text + "诊疗单据，"
            End Select
        End If
        If Len(aQryString(1)) > 0 Then .Text = .Text + "姓名：" + aQryString(1) + "，"
        If Len(aQryString(2)) > 0 And aQryString(2) <> "0" Then .Text = .Text + "性别：" + aQryString(2) + "，"
        If Len(aQryString(3)) > 0 Then .Text = .Text + "年龄：" + aQryString(3) + "～"
        If Len(aQryString(4)) > 0 Then
            If Len(aQryString(3)) = 0 Then .Text = .Text + "年龄：～"
            .Text = .Text + aQryString(4) + "，"
        Else
            If Len(aQryString(3)) > 0 Then .Text = .Text + "，"
        End If
        If Len(aQryString(5)) > 0 Then .Text = .Text + "医生：" + aQryString(5) + "，"
        If Len(aQryString(6)) > 0 Then .Text = .Text + "日期：" + aQryString(6) + "～"
        If Len(aQryString(7)) > 0 Then
            If Len(aQryString(6)) = 0 Then .Text = .Text + "日期：～"
            .Text = .Text + aQryString(7) + "，"
        End If
        If Len(aQryString(8)) > 0 Then .Text = .Text + "内容包含：" + aQryString(8) + "，"
        
        If Len(.Text) > 0 Then
            .Text = "查询条件  " + Mid(.Text, 1, Len(.Text) - 1)
        Else
            .Text = "查询条件  未设置"
        End If
    End With
End Sub

Private Sub BeginShowProgress(ByVal strCaption As String)
    With prbRefresh
        .Left = stbThis.Panels(2).Left + Me.TextWidth(strCaption) + 200
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width + stbThis.Panels(2).Left - .Left
        
        stbThis.Panels(2).Text = strCaption
        .Visible = True: Me.Refresh
    End With
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Me.mnuToolbar, 2
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

