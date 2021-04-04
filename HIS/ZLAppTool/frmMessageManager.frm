VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMessageManager 
   Caption         =   "消息收发管理"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmMessageManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Tag             =   "可变化的"
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   7695
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
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
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打开"
               Key             =   "Modify"
               Object.ToolTipText     =   "打开"
               Object.Tag             =   "打开"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "还原"
               Key             =   "Restore"
               Object.ToolTipText     =   "恢复删除"
               Object.Tag             =   "还原"
               ImageKey        =   "Restore"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "答复"
               Key             =   "Reply"
               Object.ToolTipText     =   "答复"
               Object.Tag             =   "答复"
               ImageKey        =   "Reply"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "转发"
               Key             =   "Forward"
               Object.ToolTipText     =   "转发"
               Object.Tag             =   "转发"
               ImageKey        =   "Forward"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "人员查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
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
      Top             =   5715
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   635
      SimpleText      =   $"frmMessageManager.frx":0442
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMessageManager.frx":0489
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8493
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
   Begin VB.PictureBox picCon 
      BackColor       =   &H00848484&
      FillColor       =   &H00848484&
      ForeColor       =   &H00848484&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   870
      Width           =   1485
      Begin VB.Label lblICO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已删除消息"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   330
         TabIndex        =   11
         Top             =   4230
         Width           =   900
      End
      Begin VB.Label lblICO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已发送消息"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   10
         Top             =   3090
         Width           =   900
      End
      Begin VB.Label lblICO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件箱"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   9
         Top             =   1950
         Width           =   540
      End
      Begin VB.Label lblICO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "草 稿"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   8
         Top             =   810
         Width           =   450
      End
      Begin VB.Image imgICO 
         Height          =   480
         Index           =   3
         Left            =   480
         Picture         =   "frmMessageManager.frx":0D1D
         Top             =   3630
         Width           =   480
      End
      Begin VB.Image imgICO 
         Height          =   480
         Index           =   2
         Left            =   450
         Picture         =   "frmMessageManager.frx":1027
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgICO 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "frmMessageManager.frx":1331
         Top             =   1350
         Width           =   480
      End
      Begin VB.Image imgICO 
         Height          =   480
         Index           =   0
         Left            =   450
         Picture         =   "frmMessageManager.frx":1BFB
         Top             =   240
         Width           =   480
      End
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   1485
      Left            =   2430
      TabIndex        =   6
      Top             =   3900
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2619
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMessageManager.frx":1F05
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2850
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3630
      Width           =   3000
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3480
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8684676
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":1FA2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":23F4
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":2846
            Key             =   "NewReply"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":2C98
            Key             =   "ReadReply"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":30EA
            Key             =   "Low"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":353C
            Key             =   "High"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":398E
            Key             =   "Script"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4050
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":3DE0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4220
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4440
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4660
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4880
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4A9A
            Key             =   "Reply"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4CB4
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":4ECE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":50EA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":530A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3450
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":552A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5684
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":57DE
            Key             =   "NewReply"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5938
            Key             =   "ReadReply"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5A92
            Key             =   "High"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5BEC
            Key             =   "Low"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5D46
            Key             =   "Script"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2205
      Left            =   2490
      TabIndex        =   3
      Top             =   1230
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_主题"
         Object.Tag             =   "主题"
         Text            =   "主题"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "重要性"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_发件人"
         Object.Tag             =   "发件人"
         Text            =   "发件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_收件人"
         Object.Tag             =   "收件人"
         Text            =   "收件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "时间"
         Object.Tag             =   "时间"
         Text            =   "时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   4770
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":5EA0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":60C0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":62E0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6500
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6720
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6940
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6B5A
            Key             =   "Reply"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6D74
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":6F8E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":71AA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageManager.frx":73CA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00848484&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3000
      ScaleHeight     =   405
      ScaleWidth      =   1485
      TabIndex        =   12
      Top             =   840
      Width           =   1485
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件箱"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   60
         Width           =   990
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   1770
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
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
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "还原(&S)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditReply 
         Caption         =   "答复(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditReplyAll 
         Caption         =   "全部答复(&L)"
      End
      Begin VB.Menu mnuEditForward 
         Caption         =   "转发(&W)"
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPreview 
         Caption         =   "预览窗格(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "显示已读(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLogin 
         Caption         =   "登录时有未读邮件提醒(&W)"
      End
      Begin VB.Menu mnuViewSplit6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找相关消息(&A)"
      End
      Begin VB.Menu mnuViewReflash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
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
      Begin VB.Menu mnuHelpSplit 
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
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "打开(&O)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
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
Attribute VB_Name = "frmMessageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnLoad As Boolean   '窗口还未打开时为真

Dim mstrKey As String     '未更新的邮件ID
Dim sngStartY As Single   '移动前鼠标的位置
Dim mblnItem As Boolean   '为真表示单击到ListView某一项上
Dim mintColumn As Integer '用于ListView列排序

Public mlngIndexPre As Long       '表示之前是哪个目录
Public mlngIndex As Long          '表示当前是哪个目录
Public mstrPrivs As String        '只是消息收发的模块的权限

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        
        mlngIndexPre = -1 '强制刷新
        Call FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    gblnMessageShow = True
    If gblnMessageGet = False Then
       '导航台并没有打开消息通知窗口，只有自己把它打开
       Load frmMessageRead
    End If
    Call DeleteMessage
    
    mblnLoad = True
    '-----------
    RestoreWinState Me, App.ProductName
    mnuViewShowAll.Checked = Val(zlDatabase.GetPara("显示已读邮件")) <> 0
    mnuViewLogin.Checked = Val(zlDatabase.GetPara("登录检查邮件消息")) <> 0
    
    Call 权限控制
    
    '根据LvwMain显示设置对应菜单
    mnuViewIcon_Click lvwMain.View
        
    '设置初始化选中
    mlngIndex = 1
    lblICO(mlngIndex).Tag = "下"
    '把重要性的显示位置放到第一列
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    picCon.Top = sngTop + 30
    picCon.Height = IIf(sngBottom - picCon.Top > 0, sngBottom - picCon.Top, 0)
    picCon.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = picCon.Left + picCon.Width
    
    picTitle.Top = sngTop + 30
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = picTitle.Top + picTitle.Height + 60
    
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    picTitle.Left = lvwMain.Left
    picTitle.Width = lvwMain.Width
    If rtfContent.Visible = True Then
        lvwMain.Height = (sngBottom - lvwMain.Top) * (lvwMain.Height / (lvwMain.Height + picSplitH.Height + rtfContent.Height))
        
        picSplitH.Left = lvwMain.Left
        picSplitH.Top = lvwMain.Top + lvwMain.Height
        picSplitH.Width = lvwMain.Width
        
        rtfContent.Left = lvwMain.Left
        rtfContent.Top = picSplitH.Top + picSplitH.Height
        rtfContent.Height = sngBottom - rtfContent.Top
        rtfContent.Width = lvwMain.Width
    Else
        lvwMain.Height = sngBottom - lvwMain.Top
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblnMessageShow = False
    If gblnMessageGet = False Then
        '导航台并没有打开消息通知窗口，因此退出时把它一并关了
        Unload frmMessageRead
    End If
    
    mstrKey = ""
    mlngIndexPre = 0
    Call zlDatabase.SetPara("显示已读邮件", IIf(mnuViewShowAll.Checked, 1, 0))
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Public Sub lvwMain_ItemClick(ByVal item As MSComctlLib.ListItem)
    mblnItem = True
    Call FillText
        
    SetMenu
End Sub

Private Sub ShowAttribe(ByVal strKey As String)
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim str服务对象 As String
    On Error GoTo ErrH
    gstrSQL = "select A.工作性质,A.服务对象,B.说明 from 部门性质说明 A,部门性质分类 B where A.工作性质=B.名称 and A.部门ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strKey))
    rsTemp.Close
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub
 
 Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu(1).Enabled = mnuEditNew.Enabled
        mnuShortMenu(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errHandle
    Dim strKey As String
    Dim intIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    gcnOracle.BeginTrans
    If mlngIndex <> 3 Then
        gstrSQL = "Zl_Zlmsgstate_Edit(1," & Mid(lvwMain.SelectedItem.Key, 3) & "," & lvwMain.SelectedItem.Tag & ",'" & gstrDbUser & "',Null,1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Else
        If MsgBox("你确认要删除主题为“" & lvwMain.SelectedItem.Text & "”的消息吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        Me.MousePointer = 11
        If lvwMain.SelectedItem.Tag = "0" Then
            '对于草稿，把收件人的也一并删除
            gstrSQL = "Zl_Zlmsgstate_Edit(1," & Mid(lvwMain.SelectedItem.Key, 3) & ",Null,'" & gstrDbUser & "',Null,2)"
        Else
            gstrSQL = "Zl_Zlmsgstate_Edit(1," & Mid(lvwMain.SelectedItem.Key, 3) & "," & lvwMain.SelectedItem.Tag & ",'" & gstrDbUser & "',Null,2)"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Me.MousePointer = 0
    End If
    gcnOracle.CommitTrans
    With lvwMain
        '删除ListView中对应节点
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
        Call FillText
    End With
    
    SetMenu
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    MousePointer = 0
End Sub

Private Sub mnuEditModify_Click()
    frmMessageEdit.OpenWindow Mid(lvwMain.SelectedItem.Key, 3), "", lvwMain.SelectedItem.Tag
End Sub

Private Sub mnuEditNew_Click()
    frmMessageEdit.OpenWindow "", ""
End Sub

Private Sub mnuEditReply_Click()
    frmMessageEdit.OpenWindow "", Mid(lvwMain.SelectedItem.Key, 3), lvwMain.SelectedItem.Tag, 1
End Sub

Private Sub mnuEditReplyAll_Click()
    frmMessageEdit.OpenWindow "", Mid(lvwMain.SelectedItem.Key, 3), lvwMain.SelectedItem.Tag, 2
End Sub

Private Sub mnuEditForward_Click()
    frmMessageEdit.OpenWindow "", Mid(lvwMain.SelectedItem.Key, 3), lvwMain.SelectedItem.Tag, 3
End Sub

Private Sub mnuEditRestore_Click()
'还原已删除消息
    On Error GoTo errHandle
    Dim intIndex As Long
    
    gstrSQL = "Zl_Zlmsgstate_Edit(1," & Mid(lvwMain.SelectedItem.Key, 3) & "," & lvwMain.SelectedItem.Tag & ",'" & gstrDbUser & "',Null,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With lvwMain
        '删除ListView中对应节点
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
        Call FillText
    End With
    
    SetMenu
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnuFileSaveAs_Click()
'另存为文件
    On Error Resume Next
    If rtfContent.Text = "" Then Exit Sub
    
    cdg.CancelError = True
    cdg.Filter = "RTF文件(*.RTF)|*.rtf"
    '覆盖时有提示，且不能是只读的
    cdg.flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    cdg.ShowSave
    
    If Err = 0 Then
        MousePointer = 11
        rtfContent.SaveFile cdg.FileName
        MousePointer = 0
    Else
        Err.Clear
    End If

End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewFind_Click()
    frmMessageRelate.FillList lvwMain.SelectedItem.ListSubItems(2).Tag
End Sub

Private Sub mnuViewLogin_Click()
    mnuViewLogin.Checked = Not mnuViewLogin.Checked
    Call zlDatabase.SetPara("登录检查邮件消息", IIf(mnuViewLogin.Checked, "1", "0"))
End Sub

Private Sub mnuViewPreview_Click()
    mnuViewPreview.Checked = Not mnuViewPreview.Checked
    
    picSplitH.Visible = mnuViewPreview.Checked
    rtfContent.Visible = mnuViewPreview.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewReflash_Click()
    mlngIndexPre = -1 '强制刷新
    Call FillList
End Sub

Private Sub mnuViewShowAll_Click()
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    mlngIndexPre = -1 '强制刷新
    Call FillList
End Sub

Private Sub imgICO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblICO_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgICO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblICO_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub imgICO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblICO_MouseUp(Index, Button, Shift, X, Y)
End Sub

Public Sub lblICO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Long
    For lngIndex = 0 To 3
        lblICO(lngIndex).Tag = ""
    Next
    '只有一个按钮被按下
    mlngIndex = Index
    lblICO(mlngIndex).Tag = "下"
    Call picCon_Paint
    Call FillList
End Sub

Private Sub lblICO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Long
    For lngIndex = 0 To 3
        lblICO(lngIndex).Tag = ""
    Next
    '只有一个按钮被按下
    lblICO(Index).Tag = "上"
    lblICO(mlngIndex).Tag = "下"
    Call picCon_Paint
End Sub

Private Sub lblICO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Long
    For lngIndex = 0 To 3
        lblICO(lngIndex).Tag = ""
    Next
    '只有一个按钮被按下
    lblICO(Index).Tag = "上"
    lblICO(mlngIndex).Tag = "下"
    Call picCon_Paint
End Sub

Private Sub picCon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Integer

    lngIndex = Y \ 74
    If lngIndex >= 0 And lngIndex <= 3 Then
        Call lblICO_MouseDown(lngIndex, Button, Shift, X, Y)
    End If
End Sub

Private Sub picCon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Integer
    
    If X >= 0 And X <= picCon.ScaleWidth And Y >= 0 And Y <= picCon.ScaleHeight Then
        '当进行Picture时，就捕获鼠标
        SetCapture picCon.hWnd
    Else
        '当离开Picture时，才释放鼠标
        ReleaseCapture
    End If
    lngIndex = Y \ 74
    If lngIndex >= 0 And lngIndex <= 3 And X >= 0 And X <= picCon.ScaleWidth Then
        Call lblICO_MouseMove(lngIndex, Button, Shift, X, Y)
    Else
        For lngIndex = 0 To 3
            If lngIndex <> mlngIndex Then lblICO(lngIndex).Tag = ""
        Next
        '只有一个按钮被按下
        Call picCon_Paint
    End If
    
End Sub

Private Sub picCon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Integer

    lngIndex = Y \ 74
    If lngIndex >= 0 And lngIndex <= 3 Then
        Call lblICO_MouseUp(lngIndex, Button, Shift, X, Y)
    End If
End Sub

Private Sub picCon_Paint()
    Dim rc As Rect
    Dim lngIndex As Long
    
    
    For lngIndex = 0 To 3
        rc.Left = picCon.ScaleLeft
        rc.Right = picCon.ScaleLeft + picCon.ScaleWidth
        rc.Top = lngIndex * 74
        rc.Bottom = lngIndex * 74 + 73
        
        If lblICO(lngIndex).Tag = "上" Then
            DrawEdge picCon.hDC, rc, BDR_RAISEDOUTER, BF_RECT
        ElseIf lblICO(lngIndex).Tag = "下" Then
            DrawEdge picCon.hDC, rc, BDR_SUNKENINNER, BF_RECT
        Else
            Rectangle picCon.hDC, rc.Left, rc.Top, rc.Right, rc.Bottom
        End If
    Next
End Sub

Private Sub picCon_Resize()
'对其内部控件进行排位置
'每一组留74个像素高
    Dim lngTop As Long
    Dim i As Integer
    
    For i = 0 To 3
        lngTop = i * 74
        imgICO(i).Top = lngTop + 12
        imgICO(i).Left = (picCon.ScaleWidth - imgICO(i).Width) / 2
        
        lblICO(i).Top = lngTop + 50
        lblICO(i).Left = (picCon.ScaleWidth - lblICO(i).Width) / 2
    Next
End Sub

'
Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - lvwMain.Top > 2500 And IIf(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - (sngTemp + picSplitH.Height) > 1200 Then
            picSplitH.Top = sngTemp
            lvwMain.Height = picSplitH.Top - lvwMain.Top
            rtfContent.Top = picSplitH.Top + picSplitH.Height
            rtfContent.Height = IIf(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - rtfContent.Top
        End If
        lvwMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Reply"
            mnuEditReply_Click
        Case "Forward"
            mnuEditForward_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Restore"
            mnuEditRestore_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
    End Select

End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = tlbMain.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbMain.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = tlbMain.Height
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
End Sub


Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select

End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuhelptopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub tlbMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = IIf(InStr(lblTitle.Caption, "消息") > 0, lblTitle.Caption, lblTitle.Caption & "里的消息")
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub FillList()
'功能:装入消息到lvwMain

    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    Dim strTemp As String
    Dim strICO As String
    On Error GoTo ErrH
    '如果还是同一个目录，就退出
    If mlngIndexPre = mlngIndex Then Exit Sub
    mlngIndexPre = mlngIndex
    mstrKey = ""
    '保存当前的选中项
    If Not lvwMain.SelectedItem Is Nothing Then strKey = lvwMain.SelectedItem.Key
    
    Select Case mlngIndex
        Case 0
            lblTitle.Caption = "草稿"
            gstrSQL = " S.类型=0 and S.删除=0 and S.用户=[1]" & IIf(mnuViewShowAll.Checked, "", " and substr(S.状态,1,1)='0'")
        Case 1
            lblTitle.Caption = "收件箱"
            gstrSQL = " S.类型=2 and S.删除=0 and S.用户=[1]" & IIf(mnuViewShowAll.Checked, "", " and substr(S.状态,1,1)='0'")
        Case 2
            lblTitle.Caption = "已发送消息"
            gstrSQL = " S.类型=1 and S.删除=0 and S.用户=[1]" & IIf(mnuViewShowAll.Checked, "", " and substr(S.状态,1,1)='0'")
        Case 3
            lblTitle.Caption = "已删除消息"
            gstrSQL = " S.删除=1 and S.用户=[1] " & IIf(mnuViewShowAll.Checked, "", " and substr(S.状态,1,1)='0'")
    End Select
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select M.ID,M.会话ID,M.发件人,M.收件人,M.主题,to_char(M.时间,'YYYY-MM-DD HH24:MI:SS') as 时间,S.类型,S.状态" & _
        " from zlMessages M,zlMsgState S where M.ID=S.消息ID and " & gstrSQL
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbUser)
    lvwMain.ListItems.Clear
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("状态")), "0000", rsTemp("状态"))
        
        If rsTemp("类型") = 0 Then
            strICO = "Script"
        Else
            strICO = IIf(Mid(strTemp, 1, 1) = "1", "Read", "New") & IIf(Mid(strTemp, 2, 2) <> "00", "Reply", "")   '已读+已处理
        End If
        Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("类型") & rsTemp("ID"), IIf(IsNull(rsTemp("主题")), "", rsTemp("主题")), strICO, strICO)
        If Mid(strTemp, 4, 1) <> "0" Then
            lst.SubItems(1) = IIf(Mid(strTemp, 4, 1) = 1, "高", "低")
            lst.ListSubItems(1).ReportIcon = IIf(Mid(strTemp, 4, 1) = 1, "High", "Low")
        End If
        lst.SubItems(2) = IIf(IsNull(rsTemp("发件人")), "", rsTemp("发件人"))
        lst.SubItems(3) = IIf(IsNull(rsTemp("收件人")), "", rsTemp("收件人"))
        lst.SubItems(4) = IIf(IsNull(rsTemp("时间")), "", rsTemp("时间"))
        lst.Tag = rsTemp("类型")
        lst.ListSubItems(2).Tag = rsTemp("会话ID")
        rsTemp.MoveNext
    Loop
    If lvwMain.ListItems.Count > 0 Then
        Dim item As ListItem
        On Error Resume Next
        Set item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set item = lvwMain.ListItems(1)
            item.Selected = True
            item.EnsureVisible
        Else
            Err.Clear
            item.Selected = True
            item.EnsureVisible
        End If
    End If
    '统一调用显示文本
    Call FillText
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub FillText()
'功能:把消息的内容装入到RichText中

    Dim rsTemp As New ADODB.Recordset
    
    If lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        rtfContent.Text = ""
        rtfContent.BackColor = RGB(255, 255, 255)
        Call SetMenu
        Exit Sub
    End If
    If mstrKey = lvwMain.SelectedItem.Key Then Exit Sub
    mstrKey = lvwMain.SelectedItem.Key
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 内容,背景色 from zlMessages where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(mstrKey, 3)))
    
    rtfContent.BackColor = IIf(IsNull(rsTemp("背景色")), RGB(255, 255, 255), rsTemp("背景色"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("内容")), "", rsTemp("内容"))
    Call SetMenu
End Sub

Public Sub SetMenu()
'功能:设置修改和删除按钮的有效值
    Dim blnEnabled As Boolean
    
    blnEnabled = Not (lvwMain.SelectedItem Is Nothing)
    
    tlbMain.Buttons("Modify").Enabled = blnEnabled
    tlbMain.Buttons("Delete").Enabled = blnEnabled
    mnuEditDelete.Enabled = blnEnabled
    mnuEditModify.Enabled = blnEnabled
    tlbMain.Buttons("Reply").Enabled = blnEnabled
    tlbMain.Buttons("Forward").Enabled = blnEnabled
    mnuEditReply.Enabled = blnEnabled
    mnuEditReplyAll.Enabled = blnEnabled
    mnuEditForward.Enabled = blnEnabled
    mnuViewFind.Enabled = blnEnabled
    
    mnuEditRestore.Enabled = (mlngIndex = 3 And Not (lvwMain.SelectedItem Is Nothing))
    tlbMain.Buttons("Restore").Enabled = mnuEditRestore.Enabled
    
    mnuFileSaveAs.Enabled = rtfContent.Text <> ""
    EnablePrint lvwMain.ListItems.Count > 0
    
    Dim lngCount As Long, lngSum As Long
    For lngCount = 1 To lvwMain.ListItems.Count
        If lvwMain.ListItems(lngCount).Icon = "New" Then
            lngSum = lngSum + 1
        End If
    Next
    stbThis.Panels(2).Text = "共有" & lvwMain.ListItems.Count & "条消息" & IIf(lngSum = 0, "。", "，其中有" & lngSum & "条未读。")
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'功能:设置打印和预鉴按钮的有效值
'参数:blnEnabled 有效值
    tlbMain.Buttons("Print").Enabled = blnEnabled
    tlbMain.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub DeleteMessage()
'功能：删除过时的消息
    Dim rsTemp As New ADODB.Recordset
    Dim lngDays As Long '消息能保存的天数
    
    On Error Resume Next
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 参数值 from zlOptions where 参数号=5"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub
    
    lngDays = Val(IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值")))
    If lngDays = 0 Then Exit Sub
    '删除若干天前的消息
    gstrSQL = "Zl_Zlmsgstate_Edit(2,Null,Null,Null,Null,Null,Null," & lngDays & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    mstrPrivs = GetPrivFunc(0, 12)

    If InStr(mstrPrivs, "发送消息") = 0 Then
        mnuEditNew.Enabled = False
        mnuEditNew.Visible = False
        mnuEditSplit.Visible = False
        mnuEditReply.Visible = False
        mnuEditReplyAll.Visible = False
        mnuEditForward.Visible = False
        mnuShortMenu(1).Visible = False
                
        tlbMain.Buttons("New").Visible = False
        tlbMain.Buttons("Reply").Visible = False
        tlbMain.Buttons("Forward").Visible = False
        tlbMain.Buttons("Split1").Visible = False
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

