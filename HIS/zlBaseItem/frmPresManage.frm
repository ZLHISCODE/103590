VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPresManage 
   Caption         =   "人员管理"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9645
   Icon            =   "frmPresManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPres 
      BackColor       =   &H80000005&
      Height          =   1950
      Left            =   2895
      ScaleHeight     =   1890
      ScaleWidth      =   6435
      TabIndex        =   14
      Top             =   4050
      Width           =   6495
      Begin VB.PictureBox pic说明 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   3930
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1365
         ScaleWidth      =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   420
         Width           =   45
      End
      Begin VB.PictureBox pic照片 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   1815
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1395
         ScaleWidth      =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   390
         Width           =   45
      End
      Begin VB.PictureBox pic镜框 
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   1980
         ScaleHeight     =   1380
         ScaleWidth      =   1830
         TabIndex        =   16
         Top             =   360
         Width           =   1890
         Begin VB.Image img照片 
            Height          =   1035
            Left            =   90
            Stretch         =   -1  'True
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.TextBox txt说明 
         Height          =   1455
         Left            =   4065
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2325
      End
      Begin MSComctlLib.ListView lvw人员性质_S 
         Height          =   1440
         Left            =   30
         TabIndex        =   19
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "人员性质"
            Object.Tag             =   "人员性质"
            Text            =   "人员性质"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "说明"
            Object.Tag             =   "说明"
            Text            =   "说明"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label lbl标题 
         BackColor       =   &H00808080&
         Caption         =   " 个人简介"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   4065
         TabIndex        =   22
         Top             =   45
         Width           =   2325
      End
      Begin VB.Label lbl标题 
         BackColor       =   &H00808080&
         Caption         =   " 照片"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   1965
         TabIndex        =   21
         Top             =   45
         Width           =   1905
      End
      Begin VB.Label lbl标题 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " 工作性质"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   75
         Width           =   1680
      End
   End
   Begin VB.PictureBox pic证书 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   6360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   645
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   45
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   2835
      Top             =   6045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":052A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":074A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":096A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0B8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0DAA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":0FC6
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":11E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1400
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1620
            Key             =   "start"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":183A
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1A54
            Key             =   "sign"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3810
      Top             =   6045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":232E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":254E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":276E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":298E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2BAE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2DCE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":2FEA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3204
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3424
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3644
            Key             =   "start"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":385E
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":3A78
            Key             =   "sign"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9645
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
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
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
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
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageKey        =   "start"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Object.ToolTipText     =   "停用"
               Object.Tag             =   "停用"
               ImageKey        =   "stop"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SignLine"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "证书"
               Key             =   "Sign"
               Object.ToolTipText     =   "数字证书启停用"
               Object.Tag             =   "证书"
               ImageKey        =   "sign"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SignOn"
                     Object.Tag             =   "数字证书启用"
                     Text            =   "数字证书启用"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SignOff"
                     Object.Tag             =   "数字证书停用"
                     Text            =   "数字证书停用"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sgf1"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "人员查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "人员查找"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "扩展"
               Key             =   "plugIn"
               Object.ToolTipText     =   "扩展功能"
               Object.Tag             =   "扩展"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "plugInS"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   8040
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "简码"
            Top             =   120
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   7440
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   9
            Top             =   120
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   75
               Width           =   495
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   6600
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      SimpleText      =   $"frmPresManage.frx":4752
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPresManage.frx":4799
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11933
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
   Begin MSComctlLib.ListView lvwMain 
      Height          =   1695
      Left            =   2835
      TabIndex        =   1
      Top             =   780
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   2990
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   2790
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4830
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   780
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2970
      MousePointer    =   7  'Size N S
      ScaleHeight     =   33.75
      ScaleMode       =   0  'User
      ScaleWidth      =   6165
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3945
      Width           =   6165
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   900
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":502D
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5349
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5995
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5CB1
            Key             =   "Dept_No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":5FD1
            Key             =   "Item_G"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":C833
            Key             =   "Item_W"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1575
      Top             =   1980
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
            Picture         =   "frmPresManage.frx":13095
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":136E1
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":139FD
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":13D19
            Key             =   "Dept_No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":14039
            Key             =   "Cert"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":14193
            Key             =   "Item_G"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":1A9F5
            Key             =   "Item_W"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":21257
            Key             =   "SignOn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresManage.frx":217A9
            Key             =   "SignOff"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwCert 
      Height          =   1155
      Left            =   2835
      TabIndex        =   2
      Top             =   2775
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "注册时间"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "序列号"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "使用者"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lvwLogOnOff 
      Height          =   1155
      Left            =   6480
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "StopTime"
         Text            =   "停用时间"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "StartTime"
         Text            =   "启用时间"
         Object.Width           =   3651
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcPres 
      Height          =   645
      Left            =   1065
      TabIndex        =   23
      Top             =   5805
      Width           =   1065
      _Version        =   589884
      _ExtentX        =   1879
      _ExtentY        =   1138
      _StockProps     =   64
   End
   Begin VB.Label lbl标题 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   " 数字证书"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   2850
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   2550
      Visible         =   0   'False
      Width           =   6285
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
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReport 
         Caption         =   "产生报表(&R)"
      End
      Begin VB.Menu mnuFileFile 
         Caption         =   "产生报盘文件(&F)"
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
         Caption         =   "增加人员信息(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改人员信息(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除人员信息(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPlugIn 
         Caption         =   "扩展(&E)"
         Begin VB.Menu mnuPlugItem 
            Caption         =   "功能"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "人员部门调整(&J)"
      End
      Begin VB.Menu mnuEditRole 
         Caption         =   "人员角色分配(&O)"
      End
      Begin VB.Menu mnuEditDeptRole 
         Caption         =   "批量角色分配"
      End
      Begin VB.Menu mnuEditExtend 
         Caption         =   "扩展信息维护(&E)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&T)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRegCert 
         Caption         =   "数字证书注册(&R)"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditViewCert 
         Caption         =   "查看数字证书(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImportCertPic 
         Caption         =   "导入签名图片(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDelCert 
         Caption         =   "取消证书注册"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignLine1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignOn 
         Caption         =   "数字证书启用(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSignOff 
         Caption         =   "数字证书停用(&E)"
         Visible         =   0   'False
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
      Begin VB.Menu mnuViewSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStretch 
         Caption         =   "照片自动缩放(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "人员查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewSplit6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowStopDept 
         Caption         =   "显示停用部门(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "显示停用人员(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShow 
         Caption         =   "只显示直属人员(&H)"
         Checked         =   -1  'True
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
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "扩展(&E)"
         Index           =   4
         Begin VB.Menu mnuShortPlugInItem 
            Caption         =   "功能"
            Index           =   0
         End
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "人员部门调整(&J)"
         Index           =   6
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "人员角色分配(&O)"
         Index           =   7
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
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuShortSign 
         Caption         =   "数字证书启用(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuShortSign 
         Caption         =   "数字证书停用(&E)"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPresManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Private mblnItem As Boolean  '为真表示单击到ListView某一项上
Private mintColumn As Integer
Private mblnLoad As Boolean
Private mstrKey As String    '上一次的Node关键值
Private Const mstrLvw As String = "姓名,2000,0,1;编号,800,0,2;管理职务,1400,0,0;专业技术职务,1400,0,0;住院抗菌药物权限,1400,0,0;门诊抗菌药物权限,1400,0,0;" & _
                                  "手术等级,1000,0,0;聘任技术职务,1400,0,0;出生日期,1200,0,0;身份证号,1800,0,0;性别,800,0,0;民族,800,0,0;学历,800,0,0;" & _
                                  "办公室电话,1400,0,0;移动电话,1400,0,0;电子邮件,1400,0,0;简码,600,0,0;所属部门,2000,0,0;建档时间,1440,0,0;撤档时间,1440,0,0"

Private mobjESign As Object                 '电子签名接口
Private mintCA As Integer                   '电子签名认证中心
Private mlngMode As Long
Public mstrPrivs As String                  '权限串
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset          '主界面查询
Private mstrFindValue As String             '记录查询文本框的值
Private mrsPersonProper As ADODB.Recordset  '工作性质
Private mblnCAOnOff As Boolean              '数字证书启停用权限
Private mobjForm As frmDeptExtend
Private mblnPACSInterface As Boolean        '启用影像信息系统接口
Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了正确计算coolbar的高度
        If InStr(mstrPrivs, "所有部门") = 0 Then
            Call FillTreePrivs
        Else
            If FillTree = False Then
                Unload Me
            End If
        End If
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    
    mblnLoad = True
        
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    Call InitTabControl
    
    mblnCAOnOff = InStr(mstrPrivs, ";数字证书启停用;") > 0
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    '允许进行列删除的ListView须做标记
    lvwMain.Tag = "可变化的"
    '-----------
    RestoreWinState Me, App.ProductName
    lvw人员性质_S.Visible = True
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL3_INSIDE_222_1")
    
    mnuViewShow.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "只显示直属人员", 0)) = 1)
    mnuViewStretch.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "照片自动缩放", 1)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", 0)) = 1)
    
    If InStr(1, mstrPrivs, ";修改时不限定人员性质;") = 0 Then
        gstrSQL = "Select 1 From 人员性质说明 Where 人员id =  [1] and rownum>=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "操作员是否设置权限", glngUserId)
        
        If rsTemp.RecordCount > 0 Then
            gstrSQL = "Select Distinct 人员id From 人员性质说明 Where 人员性质 In (Select 人员性质 From 人员性质说明 Where 人员id = [1])"
        Else
            '查询和当前操作员具有相同性质的人员
            gstrSQL = "Select ID As 人员id" & vbNewLine & _
                "From 人员表" & vbNewLine & _
                "Where ID Not In (Select Distinct 人员id From 人员性质说明)"
        End If
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "查询操作员工作性质", glngUserId)
    End If
    
    Call Set照片缩放
    Call Set权限控制 '包含对电子签名接口的初始化的判断
    
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
'    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
'    End If
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
     
    Call LoadPlugInMnu(Not gobjPlugIn Is Nothing)
    
    '根据LvwMain显示设置对应菜单
     mnuViewIcon_Click lvwMain.View
     
    '初始化新网RIS接口
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub InitTabControl()
    '初始化Tabcontrol控件
    With Me.tbcPres
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        
        Set mobjForm = New frmDeptExtend
        Call SetFormVisible(mobjForm.hwnd) '将窗体最大最小化隐藏

        .InsertItem(0, "人员性质", picPres.hwnd, 0).Tag = "人员性质"
        .InsertItem(1, "扩展信息", mobjForm.hwnd, 0).Tag = "扩展信息"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub LoadPlugInMnu(ByVal blnHave As Boolean)
'参数：blnHave true 表示插件对象存在
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    mnuPlugIn.Visible = blnHave
    mnuShortMenu(4).Visible = blnHave
    Toolbar1.Buttons("plugIn").Visible = blnHave
    Toolbar1.Buttons("plugInS").Visible = blnHave
 
    If blnHave Then
        'blnHave 为true 时可以确保 gobjPlugIn 对象不为 Nothing
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 GetFuncNames 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuPlugItem(i)
                Load mnuShortPlugInItem(i)
            End If
            
            mnuPlugItem(i).Caption = CStr(arrTmp(i))
            mnuPlugItem(i).Tag = CStr(arrTmp(i))
            mnuShortPlugInItem(i).Caption = CStr(arrTmp(i))
            mnuShortPlugInItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuPlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIF(i = 9, 0, i + 1) & ")"
                mnuShortPlugInItem(i).Caption = mnuPlugItem(i).Caption
            End If
        Next
    End If
End Sub

Private Sub lvwCert_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    
    Dim lngId As Long
    
    lngId = Val(Mid(lvwCert.ListItems(lvwCert.SelectedItem.Index).Key, 2))
    Call FillLogOnOff(lngId)
End Sub

Private Sub lvwCert_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        '调用弹出菜单
        If lvwCert.ListItems.Count <= 0 Then Exit Sub
        
        Dim i As Integer
        
        '隐藏其它菜单项
        For i = mnuShortIcon.LBound To mnuShortIcon.UBound
            mnuShortIcon(i).Visible = False
        Next
        For i = mnuShortMenu.LBound To mnuShortMenu.UBound
            mnuShortMenu(i).Visible = False
        Next
        For i = mnuShortSign.LBound To mnuShortSign.UBound
            mnuShortSign(i).Visible = True
        Next
        mnuShortsplit1.Visible = False
        
        mnuShortSign(0).Enabled = mnuEditSignOn.Enabled
        mnuShortSign(1).Enabled = mnuEditSignOff.Enabled
        
        '弹出菜单
        PopupMenu mnuShort
        
        '恢复其它菜单项
        For i = mnuShortIcon.LBound To mnuShortIcon.UBound
            mnuShortIcon(i).Visible = True
        Next
        For i = mnuShortMenu.LBound To mnuShortMenu.UBound
            mnuShortMenu(i).Visible = True
        Next
        For i = mnuShortSign.LBound To mnuShortSign.UBound
            mnuShortSign(i).Visible = False
        Next
        mnuShortsplit1.Visible = True
    End If
End Sub

Private Sub mnuEditExtend_Click()
    Dim strKey As String
    Dim strName As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    With lvwMain.SelectedItem
        strKey = Mid(.Key, 2)
        strName = .Text
    End With
    
    Call frmDeptExtend.ShowMe(Me, strKey, strName, 1, 1)
    Call mobjForm.initVSf(Val(strKey), 1)
End Sub

Private Sub mnuEditImportCertPic_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If Not mobjESign Is Nothing Then
        If mobjESign.RegisterCertificate(arrData) Then
            If arrData(0) <> lvwMain.SelectedItem.Text Then
                If MsgBox("该数字证书是颁发给""" & arrData(0) & """，而当前注册人员为""" & lvwMain.SelectedItem.Text & """，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            '保存签名图片
            If UBound(arrData) > 4 Then
                If arrData(5) <> "" Then
                    If SaveSignPIC(Mid(lvwMain.SelectedItem.Key, 2), arrData(5)) = False Then
                        GoTo errH
                    End If
                End If
            End If
            
            Call ShowAttribe
            Call SetMenu
            
            MsgBox lvwMain.SelectedItem.Text & "的签名图片更新成功。", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSignOff_Click()
    Dim lngId As Long
    Dim strSQL As String
    
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    If lvwCert.SelectedItem.Index <= 0 Then
        MsgBox "请选定一个数字证书！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("是否确定“停用数字签名”操作？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lngId = Val(Mid(lvwCert.SelectedItem.Key, 2))
    
    strSQL = "Zl_人员证书记录_Esignswitch(" & lngId & ",1,null)"
    Call zlDatabase.ExecuteProcedure(strSQL, "停用数字签名")
    
    'Call FillLogOnOff(lngId)
    lngId = lvwCert.SelectedItem.Index
    Call ShowAttribe
    lvwCert.ListItems(lngId).Selected = True
    Call lvwCert_ItemClick(lvwCert.SelectedItem)
End Sub

Private Sub mnuEditSignOn_Click()
    Dim lngId As Long
    Dim strSQL As String, strStop As String
    
    If lvwCert.ListItems.Count <= 0 Then Exit Sub
    If lvwCert.SelectedItem.Index <= 0 Then
        MsgBox "请选定一个数字证书！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("是否确定“启用数字签名”操作？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lngId = Val(Mid(lvwCert.SelectedItem.Key, 2))
    strStop = lvwLogOnOff.ListItems(1).Text
    
    strSQL = "Zl_人员证书记录_Esignswitch(" & lngId & ",0,to_date('" & strStop & "', 'yyyy-mm-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "启用数字签名")
    
    'Call FillLogOnOff(lngID)
    lngId = lvwCert.SelectedItem.Index
    Call ShowAttribe
    lvwCert.ListItems(lngId).Selected = True
    Call lvwCert_ItemClick(lvwCert.SelectedItem)
End Sub

Private Sub mnuPlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuPlugItem(Index).Tag)
End Sub

Private Sub mnuShortPlugInItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuShortPlugInItem(Index).Tag)
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lng人员id As Long
    
    If Not lvwMain.SelectedItem Is Nothing Then
        With lvwMain.SelectedItem
            lng人员id = Val(Mid(.Key, 2))
        End With
    End If
    
    On Error Resume Next
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, glngModul, strFunName, lng人员id, 0, 0)
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 ExecuteFunc 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim lngCert As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    cbrH = IIF(CoolBar1.Visible, CoolBar1.Height, 0)
    staH = IIF(stbThis.Visible, stbThis.Height, 0)
    lngCert = IIF(lvwCert.Visible, lvwCert.Height + lbl标题(3).Height + 30, 0)
    
    tvwMain_S.Left = 0
    tvwMain_S.Top = cbrH
    tvwMain_S.Height = Me.ScaleHeight - cbrH - staH
    
    picSplitV.Top = tvwMain_S.Top
    picSplitV.Left = tvwMain_S.Left + tvwMain_S.Width
    picSplitV.Height = tvwMain_S.Height
    
    lvwMain.Top = cbrH
    lvwMain.Left = picSplitV.Left + picSplitV.Width
    lvwMain.Width = Me.ScaleWidth - picSplitV.Left - picSplitV.Width
    lvwMain.Height = Me.ScaleHeight - cbrH - staH - lngCert - picSplitH.Height - 2000
    
    lbl标题(3).Top = lvwMain.Top + lvwMain.Height + 15
    lbl标题(3).Left = lvwMain.Left + 15
    lbl标题(3).Width = lvwMain.Width - 30
    
    lvwLogOnOff.Top = lbl标题(3).Top + lbl标题(3).Height + 15
    If mblnLoad Then lvwLogOnOff.Width = lvwMain.Width \ 2
    lvwLogOnOff.Left = lvwMain.Left + lvwMain.Width - lvwLogOnOff.Width - pic证书.Width + 30
    lvwLogOnOff.Height = lvwCert.Height
    
    pic证书.Top = lvwLogOnOff.Top
    pic证书.Left = lvwLogOnOff.Left - pic证书.Width
    pic证书.Height = lvwCert.Height
    
    lvwCert.Left = lvwMain.Left
    lvwCert.Top = lbl标题(3).Top + lbl标题(3).Height + 15
    lvwCert.Width = lvwMain.Width - lvwLogOnOff.Width - pic证书.Width
    
    picSplitH.Left = lvwMain.Left
    picSplitH.Top = lvwMain.Top + lvwMain.Height + lngCert
    picSplitH.Width = lvwMain.Width
    
    tbcPres.Move picSplitV.Left + picSplitV.Width + 30, picSplitH.Top + picSplitH.Height, Me.ScaleWidth - picSplitV.Left - picSplitV.Width - 30, Me.ScaleHeight - picSplitH.Top - picSplitH.Height - staH
    picPres.Move 0, 360, tbcPres.Width, tbcPres.Height - 360
    
    lbl标题(0).Left = 0
    lbl标题(0).Top = 0
    lbl标题(0).Width = 1800
    lvw人员性质_S.Left = 0
    lvw人员性质_S.Top = lbl标题(0).Top + lbl标题(0).Height + 15
    lvw人员性质_S.Width = lbl标题(0).Width
    lvw人员性质_S.Height = picPres.ScaleHeight - lvw人员性质_S.Top
    
    pic照片.Left = lvw人员性质_S.Left + lvw人员性质_S.Width
    pic照片.Top = lbl标题(0).Top
    pic照片.Height = picPres.ScaleHeight
    
    lbl标题(1).Top = lbl标题(0).Top
    lbl标题(1).Left = pic照片.Left + pic照片.Width
    lbl标题(1).Width = 1800
    pic镜框.Left = lbl标题(1).Left
    pic镜框.Top = lvw人员性质_S.Top
    pic镜框.Width = lbl标题(1).Width
    pic镜框.Height = lvw人员性质_S.Height
    
    pic说明.Top = pic照片.Top
    pic说明.Left = pic镜框.Left + pic镜框.Width
    pic说明.Height = pic照片.Height
    
    lbl标题(2).Top = lbl标题(0).Top
    lbl标题(2).Left = pic说明.Left + pic说明.Width
    lbl标题(2).Width = picPres.ScaleWidth - lbl标题(2).Left
    
    txt说明.Left = lbl标题(2).Left
    txt说明.Top = lvw人员性质_S.Top
    txt说明.Height = lvw人员性质_S.Height
    txt说明.Width = lbl标题(2).Width
    
    SetParent txtFind.hwnd, Toolbar1.hwnd
    SetParent picFind.hwnd, Toolbar1.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "只显示直属人员", IIF(mnuViewShow.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", IIF(mnuViewShowStop.Checked, 1, 0)
        
    SaveWinState Me, App.ProductName
    
    Set mobjESign = Nothing
    If Not mobjForm Is Nothing Then Set mobjForm = Nothing
    
End Sub

Private Sub img照片_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        msngStartY = Y
    End If
End Sub

Private Sub img照片_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngLeft As Single
    Dim sngTop As Single
    
    '缩放状态不处理
    If mnuViewStretch.Checked = True Then Exit Sub
    If Button = 1 Then
        '首先求出可能的
        sngLeft = img照片.Left + X - msngStartX
        sngTop = img照片.Top + Y - msngStartY
        
        '设置可能的左边距
        If img照片.Width < pic镜框.ScaleWidth Or sngLeft > pic镜框.ScaleLeft Then
            sngLeft = pic镜框.ScaleLeft
        Else
            If sngLeft + img照片.Width < pic镜框.ScaleWidth Then
                sngLeft = pic镜框.ScaleWidth - img照片.Width
            End If
        End If
        '设置可能的顶边距
        If img照片.Height < pic镜框.ScaleHeight Or sngTop > pic镜框.ScaleTop Then
            sngTop = pic镜框.ScaleTop
        Else
            If sngTop + img照片.Height < pic镜框.ScaleHeight Then
                sngTop = pic镜框.ScaleHeight - img照片.Height
            End If
        End If
        img照片.Left = sngLeft
        img照片.Top = sngTop
    End If
End Sub

Private Sub lbl标题_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Index = 3 Then msngStartY = Y
End Sub

Private Sub lbl标题_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Index = 3 Then
        If lvwMain.Height + Y - msngStartY < 2000 Or lvwCert.Height - (Y - msngStartY) < 500 Then Exit Sub
        lbl标题(3).Top = lbl标题(3).Top + Y - msngStartY
        lvwMain.Height = lvwMain.Height + Y - msngStartY
        lvwCert.Top = lvwCert.Top + Y - msngStartY
        lvwCert.Height = lvwCert.Height - (Y - msngStartY)
        lvwLogOnOff.Top = lvwCert.Top
        lvwLogOnOff.Height = lvwCert.Height
        pic证书.Top = lvwCert.Top
        pic证书.Height = lvwLogOnOff.Height
    End If
End Sub

Private Sub lvwCert_DblClick()
    Call lvwCert_KeyPress(13)
End Sub

Private Sub lvwCert_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwCert.SelectedItem Is Nothing Then
            If mnuEditViewCert.Enabled And mnuEditViewCert.Visible Then
                Call mnuEditViewCert_Click
            End If
        End If
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    SetMenu
End Sub

Public Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowAttribe
    Call mobjForm.initVSf(Val(Mid(Item.Key, 2)), 1)
    
    mblnItem = True
    
    Call lvwMain_GotFocus
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu(1).Enabled = mnuEditNew.Enabled
        mnuShortMenu(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu(3).Enabled = mnuEditDelete.Enabled
        mnuShortMenu(5).Enabled = mnuEditAdjust.Enabled
        mnuShortMenu(6).Enabled = mnuEditRole.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditAdjust_Click()
    '人员部门调整
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Mid(lvwMain.SelectedItem.Key, 2) = glngUserId And InStr(mstrPrivs, "所有部门") = 0 Then
        MsgBox "不允许对当前登录人员进行“人员部门调整”！", vbInformation, gstrSysName
        Exit Sub
    End If
    With frmPresAdjust
        .EntryPort Mid(lvwMain.SelectedItem.Key, 2), mstrPrivs
        .Show vbModal, Me
        mstrKey = ""
        Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    End With
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Integer
    Dim blnRisTrans As Boolean
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "所有部门") = 0 Then
        If glngUserId = Val(Mid(lvwMain.SelectedItem.Key, 2)) Then
            MsgBox "不能删除当前登录用户对应的人员！", vbInformation, gstrSysName
            Exit Sub
        End If
        'If CheckDeptPermission(glngUserId, Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
    End If
    
    If MsgBox("你确认要删除姓名为“" & lvwMain.SelectedItem.Text & "”的人员吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo ErrHandle
        
        '新网RIS接口，删除人员信息；标准版，启用参数，部门性质为“检查”的部门人员，接口部件有效的前提下
        If Int(glngSys / 100) = 1 And mblnPACSInterface = True Then
            If IsCheckDeptPres(Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.Personnel, RISBaseItemOper.Delete, Val(Mid(lvwMain.SelectedItem.Key, 2))) <> 1 Then
                        '出错时提示接口错误信息
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                        End If

                        Exit Sub
                    End If
                    
                    blnRisTrans = True
                Else
                    '接口部件无效时禁止并提示
                    MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                    
                    Exit Sub
                End If
            End If
        End If
        
        gstrSQL = "zl_人员表_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call SQLTest
        
        blnRisTrans = False
        
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
        End With
        
        Call ShowAttribe
        If lvwMain.SelectedItem Is Nothing Then
            Call mobjForm.initVSf(0, 1)
        Else
            Call mobjForm.initVSf(Val(Mid(lvwMain.SelectedItem.Key, 2)), 1)
        End If
        Call SetMenu
    End If
    Exit Sub
ErrHandle:
    'Ris接口和HIS不同步时，写错误日志
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS删除人员信息错误，RIS接口和HIS数据不同步，请与系统管理员联系。", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmPresManage：mnuEditDelete_Click", "HIS删除人员信息错误，RIS接口和HIS数据不同步", "人员ID=" & Val(Mid(lvwMain.SelectedItem.Key, 2)), 0)
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsCheckDeptPres(ByVal lngPres As Long) As Boolean
    '是否检查科室人员
    Dim rsData  As ADODB.Recordset
    
    gstrSQL = "Select 1 From 部门人员 A, 部门性质说明 B Where a.部门id = b.部门id And 工作性质 = '检查' And a.人员id = [1] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsCheckDeptPres", lngPres)
    
    IsCheckDeptPres = Not rsData.EOF
End Function
Private Sub mnuEditDeptRole_Click()
    Dim frmTmp As frmPresRoleBat
    If tvwMain_S.SelectedItem.Key = "Root" Then Exit Sub
    Set frmTmp = New frmPresRoleBat
    frmTmp.ShowMe Me, Val(Mid(tvwMain_S.SelectedItem.Key, 2)), tvwMain_S.SelectedItem.Text
End Sub

Private Sub mnuEditModify_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    With lvwMain.SelectedItem
'        If InStr(mstrPrivs, "所有部门") = 0 Then
'            If CheckDeptPermission(glngUserId, Val(Mid(.Key, 2))) Then
               frmPresSet.编辑人员 Mid(.Key, 2)
'            End If
'        Else
'            frmPresSet.编辑人员 Mid(.Key, 2)
'        End If

        If InStr(1, mstrPrivs, ";修改时不限定人员性质;") = 0 Then
        gstrSQL = "Select 1 From 人员性质说明 Where 人员id =  [1] and rownum>=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "操作员是否设置权限", glngUserId)
        
        If rsTemp.RecordCount > 0 Then
            gstrSQL = "Select Distinct 人员id From 人员性质说明 Where 人员性质 In (Select 人员性质 From 人员性质说明 Where 人员id = [1])"
        Else
            '查询和当前操作员具有相同性质的人员
            gstrSQL = "Select ID As 人员id" & vbNewLine & _
                "From 人员表" & vbNewLine & _
                "Where ID Not In (Select Distinct 人员id From 人员性质说明)"
        End If
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "查询操作员工作性质", glngUserId)
    End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditNew_Click()
    On Error GoTo ErrHandle
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    With tvwMain_S.SelectedItem
        If InStr(mstrPrivs, "所有部门") = 0 And .ForeColor <> vbBlack Then
            MsgBox "你不能在“" & .Text & "”下增加人员信息！", vbInformation, gstrSysName
            Exit Sub
        End If
        frmPresSet.编辑人员 , Mid(.Key, 2)
        
        '查询和当前操作员具有相同性质的人员
        gstrSQL = "Select ID As 人员id" & vbNewLine & _
                "From 人员表" & vbNewLine & _
                "Where ID Not In (Select Distinct 人员id From 人员性质说明)" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select Distinct 人员id From 人员性质说明 Where 人员性质 In(Select 人员性质 From 人员性质说明 Where 人员id = [1])"
        Set mrsPersonProper = zlDatabase.OpenSQLRecord(gstrSQL, "查询操作员工作性质", glngUserId)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditRegCert_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If Not mobjESign Is Nothing Then
        If mobjESign.RegisterCertificate(arrData, Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
            If arrData(0) <> lvwMain.SelectedItem.Text Then
                If MsgBox("该数字证书是颁发给""" & arrData(0) & """，而当前注册人员为""" & lvwMain.SelectedItem.Text & """，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            '保存签名图片
            gcnOracle.BeginTrans
            If UBound(arrData) > 4 Then
                If arrData(5) <> "" Then
                    If SaveSignPIC(Mid(lvwMain.SelectedItem.Key, 2), arrData(5)) = False Then
                        GoTo errH
                    End If
                End If
            End If
            
            gstrSQL = "zl_人员证书记录_Insert(" & _
                Val(Mid(lvwMain.SelectedItem.Key, 2)) & "," & _
                "'" & Replace(arrData(1), "'", "''") & "'," & _
                "'" & Replace(arrData(2), "'", "''") & "'," & _
                "'" & Replace(arrData(3), "'", "''") & "'," & _
                "'" & Replace(arrData(4), "'", "''") & "'," & _
                "'" & Replace(arrData(6), "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If CStr(arrData(7)) <> "" Then
                If Not Sys.SaveLob(glngSys, 14, Val(Mid(lvwMain.SelectedItem.Key, 2)) & "," & Trim(arrData(2)), CStr(arrData(7)), 1) Then
                    GoTo errH
                End If
            End If
            Call ShowAttribe
            Call SetMenu
            gcnOracle.CommitTrans
            
            MsgBox "数字证书注册成功，""" & lvwMain.SelectedItem.Text & """可以在其他场合使用该证书进行电子签名。", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelCert_Click()
    Dim arrData As Variant
    
    On Error GoTo errH
    
    If MsgBox("确实要取消人员""" & lvwMain.SelectedItem.Text & """当前选择的数字证书注册吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "zl_人员证书记录_Delete(" & Val(Mid(lvwCert.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call ShowAttribe
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditRole_Click()
    Dim frmTmp As frmPresRole
    '人员角色分配
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Mid(lvwMain.SelectedItem.Key, 2) = glngUserId Then
        MsgBox "不允许对当前登录人员进行“人员角色分配”！", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckIsUser(Mid(lvwMain.SelectedItem.Key, 2)) = False Then
        Exit Sub
    End If
    Set frmTmp = New frmPresRole
    Call frmTmp.ShowMe(Me, Val(Mid(lvwMain.SelectedItem.Key, 2)))
End Sub

Private Sub mnuEditStart_Click()
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If tvwMain_S.SelectedItem.Image = "Dept_No" Then
        MsgBox "该人员所属部门还是停用状态，请到部门管理中启用对于部门！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "Zl_人员表_启用(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
            
    '执行启用过程
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mstrKey = ""
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
'    If InStr(mstrPrivs, "所有部门") = 0 Then
'        Call FillTreePrivs
'    Else
'        Call FillTree
'    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub mnuEditStop_Click()
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "所有部门") = 0 And Mid(lvwMain.SelectedItem.Key, 2) = glngUserId Then
        MsgBox "拒绝停用用户对应的人员！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmPresStop.编辑人员 (Mid(lvwMain.SelectedItem.Key, 2))
    mstrKey = ""
    Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    
'    If InStr(mstrPrivs, "所有部门") = 0 Then
'        Call FillTreePrivs
'    Else
'        Call FillTree
'    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditViewCert_Click()
    If Not mobjESign Is Nothing Then
        If mobjESign.ViewCertificate(Val(Mid(lvwCert.SelectedItem.Key, 2))) Then
            
        End If
    End If
End Sub

Private Sub mnuFileFile_Click()
'    Call 调查报盘(Me)
    Dim objNow As Object
    
    On Error Resume Next
    
    Set objNow = CreateObject("zl9MedRec.ClsMedRec")
    Call objNow.调查报盘(Me)
End Sub

Private Sub mnuFileReport_Click()
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    On Error Resume Next
    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1002_1", Me, "人员ID=" & Mid(lvwMain.SelectedItem.Key, 2), 1
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：部门=部门id，人员=人员id
    Dim lng部门ID As Long
    Dim lng人员id As Long
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        If tvwMain_S.SelectedItem.Key <> "Root" Then
            lng部门ID = Mid(tvwMain_S.SelectedItem.Key, 2)
        End If
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng人员id = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "部门=" & IIF(lng部门ID = 0, "", lng部门ID), _
        "人员=" & IIF(lng人员id = 0, "", lng人员id))
End Sub

Private Sub mnuShortSign_Click(Index As Integer)
    Select Case Index
        Case 0      '数字证书启用
            Call mnuEditSignOn_Click
        Case 1      '数字证书停用
            Call mnuEditSignOff_Click
    End Select
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '列有变化就要重新刷新
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewFind_Click()
    frmPresFind.ShowOfType Me, 0, mnuViewShowStop.Checked
End Sub

Private Sub mnuViewReflash_Click()
    If InStr(mstrPrivs, "所有部门") = 0 Then
        FillTreePrivs
    Else
        FillTree
    End If
End Sub

Private Sub mnuViewShow_Click()
    mnuViewShow.Checked = Not mnuViewShow.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStopDept_Click()
    mnuViewShowStopDept.Checked = Not mnuViewShowStopDept.Checked
    Call FillTree
End Sub

Private Sub mnuViewStretch_Click()
    mnuViewStretch.Checked = Not mnuViewStretch.Checked
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "照片自动缩放", IIF(mnuViewStretch.Checked, 1, 0)
    Call Set照片缩放
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If tbcPres.Width - X < 2000 Or tvwMain_S.Width + X < 2000 Then Exit Sub
        picSplitV.Left = picSplitV.Left + X
        tvwMain_S.Width = tvwMain_S.Width + X
        lvwMain.Left = lvwMain.Left + X
        lvwMain.Width = lvwMain.Width - X
        
        lbl标题(3).Left = lbl标题(3).Left + X
        lbl标题(3).Width = lbl标题(3).Width - X
        lvwCert.Left = lvwCert.Left + X
        lvwCert.Width = lvwCert.Width - X
        
        picSplitH.Left = picSplitH.Left + X
        picSplitH.Width = picSplitH.Width - X
        
        tbcPres.Left = tbcPres.Left + X
        tbcPres.Width = tbcPres.Width - X
        picPres.Width = picPres.Width - X
        
        lbl标题(2).Width = lbl标题(2).Width - X
        txt说明.Width = txt说明.Width - X
        
        tvwMain_S.SetFocus
    End If
End Sub

Private Sub pic镜框_Resize()
    If mnuViewStretch.Checked = True Then
        '缩放
        img照片.Width = pic镜框.ScaleWidth
        img照片.Height = pic镜框.ScaleHeight
    Else
        '处理不完整
        If pic镜框.ScaleWidth > img照片.Width Then
            img照片.Left = pic镜框.ScaleLeft
        Else
            If img照片.Left + img照片.Width < pic镜框.ScaleWidth Then
                img照片.Left = pic镜框.ScaleWidth - img照片.Width
            End If
        End If
        
        If pic镜框.ScaleHeight > img照片.Height Then
            img照片.Top = pic镜框.ScaleTop
        Else
            If img照片.Top + img照片.Height < pic镜框.ScaleHeight Then
                img照片.Top = pic镜框.ScaleHeight - img照片.Height
            End If
        End If
    End If
End Sub

Private Sub pic照片_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic照片_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lvw人员性质_S.Width + X < 1000 Or pic镜框.Width - X < 1000 Then Exit Sub
        
        lvw人员性质_S.Width = lvw人员性质_S.Width + X
        lbl标题(0).Width = lvw人员性质_S.Width
        
        pic照片.Left = pic照片.Left + X
        
        pic镜框.Left = pic镜框.Left + X
        pic镜框.Width = pic镜框.Width - X
        lbl标题(1).Left = pic镜框.Left
        lbl标题(1).Width = pic镜框.Width
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub pic说明_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic说明_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If pic镜框.Width + X < 1000 Or txt说明.Width - X < 1000 Then Exit Sub
        
        pic镜框.Width = pic镜框.Width + X
        lbl标题(1).Width = pic镜框.Width
        
        pic说明.Left = pic说明.Left + X
        
        txt说明.Left = txt说明.Left + X
        txt说明.Width = txt说明.Width - X
        lbl标题(2).Left = txt说明.Left
        lbl标题(2).Width = txt说明.Width
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lvwCert.Visible Then
            If tbcPres.Height - Y < 800 Or lvwCert.Height + Y < 500 Then Exit Sub
        Else
            If tbcPres.Height - Y < 800 Or lvwMain.Height + Y < 2000 Then Exit Sub
        End If
        
        picSplitH.Top = picSplitH.Top + Y
                
        If lvwCert.Visible Then
            lvwCert.Height = lvwCert.Height + Y
            pic证书.Height = lvwCert.Height
            lvwLogOnOff.Height = lvwCert.Height
        Else
            lvwMain.Height = lvwMain.Height + Y
        End If
        
        tbcPres.Top = tbcPres.Top + Y
        tbcPres.Height = tbcPres.Height - Y
        picPres.Height = picPres.Height - Y

        lvw人员性质_S.Height = lvw人员性质_S.Height - Y

        pic镜框.Height = pic镜框.Height - Y
        txt说明.Height = txt说明.Height - Y
        pic照片.Height = pic照片.Height - Y
        pic说明.Height = pic说明.Height - Y
        
        lvwMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub pic证书_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub pic证书_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = pic证书.Left + X - msngStartX
        If sngTemp - lvwCert.Left > 1000 And ScaleWidth - (sngTemp + pic证书.Width) > 1000 Then
            pic证书.Left = sngTemp
            lvwCert.Width = pic证书.Left - lvwCert.Left
            lvwLogOnOff.Left = pic证书.Left + pic证书.Width
            lvwLogOnOff.Width = lvwMain.Left + lvwMain.Width - pic证书.Width - lvwCert.Width - lvwCert.Left
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "Find"
            mnuViewFind_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "plugIn"
            PopupMenu mnuPlugIn, vbPopupMenuRightButton
        Case "Sign"
            '不在此处处理
    End Select
End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Sign" Then
        Button.ButtonMenus("SignOn").Enabled = mnuEditSignOn.Enabled
        Button.ButtonMenus("SignOff").Enabled = mnuEditSignOff.Enabled
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    If ButtonMenu.Key = "SignOn" Then
        Call mnuEditSignOn_Click
    ElseIf ButtonMenu.Key = "SignOff" Then
        Call mnuEditSignOff_Click
    Else
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
            Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "√", "  ")
        Next
        mnuViewIcon(ButtonMenu.Index - 1).Checked = True
        Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text, "  ", "√")
        lvwMain.View = ButtonMenu.Index - 1
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    Me.mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
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
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "√", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "√")
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
        Case 6
            mnuEditAdjust_Click
        Case 7
            mnuEditRole_Click
    End Select
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuhelptopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwMain_S_GotFocus()
    SetMenu
    If mnuViewShow.Checked = True Then
        stbThis.Panels(2).Text = "该部门有人员" & lvwMain.ListItems.Count & "名（包括下级部门）。"
    Else
        stbThis.Panels(2).Text = "该部门有人员" & lvwMain.ListItems.Count & "名。"
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim objNode As Node
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Set objNode = tvwMain_S.SelectedItem
    If mstrKey = objNode.Key Then Exit Sub
    mstrKey = objNode.Key
    
    If objNode.ForeColor = &H8000000C Then
        lvwMain.ListItems.Clear
    Else
        FillList objNode.Key
    End If
    If mnuViewShow.Checked = True Then
        stbThis.Panels(2).Text = "该部门有人员" & lvwMain.ListItems.Count & "名（包括下级部门）。"
    Else
        stbThis.Panels(2).Text = "该部门有人员" & lvwMain.ListItems.Count & "名。"
    End If
    SetMenu
End Sub


Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    Dim str单位 As String
    
    str单位 = GetUnitName
    objPrint.Title.Text = str单位 & "人员表"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(Sys.Currentdate, "yyyy年MM月dd日")
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

Private Sub FillTreePrivs()
'功能:装入所属部门到tvwMain_S
    Dim nodTmp As Node
    Dim rsDeptID As ADODB.Recordset
    Dim strTemp As String
    strTemp = "Dept"
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Max(Level) as 层,A.ID,A.上级ID,A.名称,'【'||A.编码||'】' 编码 " & _
              "From 部门表 A Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior 上级ID=ID " & _
              "Group by A.ID,A.上级ID,A.名称,A.编码 " & _
              "Order by A.编码,层 Desc"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvwMain_S
        .Sorted = True
        .Nodes.Clear
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!上级id), 0, rsDeptID!上级id) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!上级id, tvwChild, "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.ForeColor = &H8000000C
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    '生成子结点
    gstrSQL = "Select ID,上级ID,'【'||编码||'】' 编码,名称 " & _
              "From 部门表 A " & _
              "Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior ID=上级ID"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvwMain_S
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!上级id), 0, rsDeptID!上级id) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!上级id, tvwChild, "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.ForeColor = vbBlack
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    
    If tvwMain_S.Nodes.Count > 0 Then tvwMain_S.Nodes(1).Selected = True
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FillTree() As Boolean
'功能:装入所有部门到tvwMain_S
'参数:
    Dim strTemp As String
    Dim strKey As String
    Dim rs部门 As New ADODB.Recordset
    
    
    mstrKey = ""
    
    On Error GoTo ErrHandle
    rs部门.CursorLocation = adUseClient
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    If mnuViewShowStopDept.Checked = True Then
        strTemp = ""
    Else
        strTemp = " where (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null ) "
    End If
    
    gstrSQL = "select id,上级id,编码 ,名称,to_char(撤档时间,'YYYY-MM-DD') as 撤档时间  from 部门表 " & strTemp & " start with 上级id is null connect by prior id =上级id "
    Set rs部门 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    tvwMain_S.Nodes.Clear
    tvwMain_S.Nodes.Add , , "Root", "所有部门", "Root", "Root"
    tvwMain_S.Nodes("Root").Sorted = True
'    strTemp = "Dept"
    Do Until rs部门.EOF
        If CDate(IIF(IsNull(rs部门("撤档时间")), CDate("3000/1/1"), rs部门("撤档时间"))) = CDate("3000/1/1") Then
            strTemp = "Dept"
        Else
            strTemp = "Dept_No"
        End If
        If IsNull(rs部门("上级id")) Then
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
        Else
            tvwMain_S.Nodes.Add "C" & rs部门("上级id"), tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
        End If
        tvwMain_S.Nodes("C" & rs部门("id")).Sorted = True
        rs部门.MoveNext
    Loop
    If tvwMain_S.Nodes.Count = 1 Then
        MsgBox "部门信息不全，不能使用本模块。" & vbCrLf & "部门信息可在“部门管理”建立。", vbInformation, gstrSysName
        FillTree = False
        Exit Function
    End If
    
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(2) '避开根节点
        nod.Selected = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        If nod.Key = "Root" Then nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    FillTree = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FillList(ByVal str部门ID As String)
'功能:装入对应部门的人员到lvwMain
'参数:str部门ID 部门的标识

    Dim rs人员 As New ADODB.Recordset
    Dim lst As ListItem
    Dim i As Integer, varValue As Variant
    Dim strKey As String
    Dim stroldkey As String
    Dim rsTemp As ADODB.Recordset
    Dim bln忽略性质 As Boolean
    
    On Error GoTo ErrHandle
    If Not lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain.SelectedItem.Key
    End If
    rs人员.CursorLocation = adUseClient
    rs人员.CursorType = adOpenKeyset
    rs人员.LockType = adLockReadOnly
    
    Call FS.ShowFlash("正在获取人员信息数据,请稍候 ...", Me)
    If str部门ID = "Root" Then
        gstrSQL = "select a.ID,C.部门ID,a.姓名,a.编号,to_char(A.出生日期,'yyyy-MM-dd') as 出生日期,A.身份证号,A.性别,A.民族,a.简码 ,b.名称 as 所属部门 " & _
                    "   ,A.个人简介,A.专业技术职务,A.管理职务,Decode(D.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 住院抗菌药物权限" & _
                    "   ,Decode(D.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 门诊抗菌药物权限 " & _
                    "   ,A.手术等级,A.办公室电话,A.移动电话,A.电子邮件,A.学历,decode(A.聘任技术职务,1,'正高',2,'副高',3,'中级',4,'助理/师级',5,'员/士',9,'待聘') as 聘任技术职务" & _
                    "   ,to_char(A.建档时间,'YYYY-MM-DD') as 建档时间,Nvl(To_Char(A.撤档时间, 'YYYY-MM-DD'), '3000-01-01') as 撤档时间" & _
                    " from 人员表 a,部门表 b,部门人员 C,人员抗菌药物权限 D, 人员抗菌药物权限 E " & _
                    " where a.id = C.人员id  and C.部门ID=B.ID And D.人员ID(+)=a.id And (D.记录状态=1 or D.记录状态 is null) and d.场合(+) = 1 " & _
                    "   and e.人员id(+)=a.id and (e.记录状态=1 or e.记录状态 is null) and e.场合(+) = 2 " & _
                    "   and B.ID in (select ID from 部门表 start with 上级ID is null connect by prior id=上级ID)"
    Else
        If mnuViewShow.Checked = True Then
            gstrSQL = "select a.ID,C.部门ID,a.姓名,a.编号,to_char(A.出生日期,'yyyy-MM-dd') as 出生日期,A.身份证号,A.性别,A.民族,a.简码 ,b.名称 as 所属部门 " & _
                        " ,A.个人简介,A.专业技术职务,A.管理职务,Decode(F.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 住院抗菌药物权限" & _
                        " ,Decode(g.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 门诊抗菌药物权限" & _
                        " ,A.手术等级,A.办公室电话,A.移动电话,A.电子邮件,A.学历,decode(A.聘任技术职务,1,'正高',2,'副高',3,'中级',4,'助理/师级',5,'员/士',9,'待聘') as 聘任技术职务" & _
                        " ,to_char(A.建档时间,'YYYY-MM-DD') as 建档时间,Nvl(To_Char(A.撤档时间, 'YYYY-MM-DD'), '3000-01-01') as 撤档时间" & _
                        " from 人员表 a,部门表 b,部门人员 C, " & _
                        "(select distinct 人员ID From 部门人员 where 部门id =[1]) D,人员抗菌药物权限 F, 人员抗菌药物权限 G  " & _
                        " where A.id = C.人员id and C.部门ID=B.ID and A.ID=D.人员ID And F.人员ID(+)=a.id And g.人员ID(+)=a.id " & _
                        "   And (F.记录状态=1 or F.记录状态 is null) and f.场合(+)=1 " & _
                        "   And (g.记录状态=1 or g.记录状态 is null) and g.场合(+)=2 "
        Else
            gstrSQL = "select a.ID,C.部门ID,a.姓名,a.编号,to_char(A.出生日期,'yyyy-MM-dd') as 出生日期,A.身份证号,A.性别,A.民族,a.简码 ,b.名称 as 所属部门 " & _
                        " ,A.个人简介,A.专业技术职务,A.管理职务,Decode(F.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 住院抗菌药物权限" & _
                        " ,Decode(g.级别,1,'非限制使用',2,'限制使用',3,'特殊使用','') as 门诊抗菌药物权限 " & _
                        " ,A.手术等级,A.办公室电话,A.移动电话,A.电子邮件,A.学历,decode(A.聘任技术职务,1,'正高',2,'副高',3,'中级',4,'助理/师级',5,'员/士',9,'待聘') as 聘任技术职务" & _
                        " ,to_char(A.建档时间,'YYYY-MM-DD') as 建档时间,Nvl(To_Char(A.撤档时间, 'YYYY-MM-DD'), '3000-01-01') as 撤档时间" & _
                        " from 人员表 a,部门表 b,部门人员 C, " & _
                        "(select distinct 人员ID From 部门人员 where 部门id in " & _
                        "(select distinct id  From 部门表 start with ID=[1] connect by prior id=上级id)) D " & _
                        "  ,人员抗菌药物权限 F, 人员抗菌药物权限 G " & _
                        " where A.id = C.人员id and C.部门ID=B.ID and A.ID=D.人员ID And F.人员ID(+)=a.id And (f.记录状态=1 or f.记录状态 is null) and f.场合(+)=1 " & _
                        "   and g.人员id(+)=a.id and (g.记录状态=1 or g.记录状态 is null) and g.场合(+)=2 "
        End If
    End If
    
    If mnuViewShowStop.Checked = False Then
        gstrSQL = gstrSQL & " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
    End If
    
    gstrSQL = gstrSQL & " order by a.id"
    Set rs人员 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(str部门ID, 2)))
        
    Dim lng所属部门 As Long
    For i = 2 To lvwMain.ColumnHeaders.Count
        If lvwMain.ColumnHeaders(i).Text = "所属部门" Then lng所属部门 = i
    Next
    
    zlControl.FormLock lvwMain.hwnd
    lvwMain.ListItems.Clear
        
    Do Until rs人员.EOF
        bln忽略性质 = False
        If InStr(1, mstrPrivs, ";修改时不限定人员性质;") = 0 Then
            mrsPersonProper.MoveFirst
            Do Until mrsPersonProper.EOF
                If rs人员!ID = mrsPersonProper!人员ID Then
                    bln忽略性质 = True
                    Exit Do
                Else
                    mrsPersonProper.MoveNext
                End If
            Loop
        Else
            bln忽略性质 = True
        End If
        If bln忽略性质 = True Then
            If stroldkey <> "C" & rs人员("ID") Then
                stroldkey = "C" & rs人员("ID")
                Set lst = lvwMain.ListItems.Add(, "C" & rs人员("ID"), rs人员("姓名"), IIF(rs人员!性别 = "男", "Item", IIF(rs人员!性别 = "女", "Item_G", "Item_W")), IIF(rs人员!性别 = "男", "Item", IIF(rs人员!性别 = "女", "Item_G", "Item_W")))
                lst.Tag = IIF(IsNull(rs人员("个人简介")), "", rs人员("个人简介"))
    
                For i = 2 To lvwMain.ColumnHeaders.Count
                    varValue = rs人员(lvwMain.ColumnHeaders(i).Text).value
                    lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                Next
                
                If Format(rs人员!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    lst.ForeColor = &HFF&
                    For i = 1 To Me.lvwMain.ColumnHeaders.Count - 1
                        lst.ListSubItems(i).ForeColor = &HFF&
                    Next
                End If
            Else
                '该人员已经加入
                If lng所属部门 > 1 Then
                    '如果该列显示，那就追加到最后
                    lvwMain.ListItems("C" & rs人员("ID")).SubItems(lng所属部门 - 1) = lvwMain.ListItems("C" & rs人员("ID")).SubItems(lng所属部门 - 1) & "," & rs人员("所属部门")
                End If
                Err.Clear
            End If
        End If
            
        rs人员.MoveNext
    Loop
    zlControl.FormLock 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        Err.Clear
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
        Else
            Err.Clear
            Item.Selected = True
            Item.EnsureVisible
        End If
    End If
    Call ShowAttribe
    If lvwMain.SelectedItem Is Nothing Then
        Call mobjForm.initVSf(0, 1)
    Else
        Call mobjForm.initVSf(Val(Mid(lvwMain.SelectedItem.Key, 2)), 1)
    End If
    Call SetMenu
    Call FS.StopFlash
    
    Exit Sub

ErrHandle:
    Call FS.StopFlash
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowAttribe()
    Dim rsTemp As New ADODB.Recordset
    Dim strTempFile As String
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    lvwCert.ListItems.Clear
    lvwCert.Sorted = False
    lvw人员性质_S.ListItems.Clear
    Set img照片.Picture = Nothing
    img照片.ToolTipText = "无照片"
    txt说明.Text = ""
    
    If lvwMain.SelectedItem Is Nothing Then
        '如果没有选择中项，就清空性质列表
        Exit Sub
    End If
    rsTemp.CursorLocation = adUseClient
    
    '显示人员的证书记录
    gstrSQL = "Select ID,CertDN,CertSN,SignCert,EncCert,注册时间,是否停用 From 人员证书记录" & _
        " Where 人员ID=[1] Order by 注册时间 Desc,ID Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    
    Do Until rsTemp.EOF
        Set ObjItem = lvwCert.ListItems.Add(, "_" & rsTemp!ID, Format(rsTemp!注册时间, "yyyy-MM-dd HH:mm:ss"), , IIF(NVL(rsTemp!是否停用, 0) = 0, "SignOn", "SignOff"))
        ObjItem.SubItems(1) = "" & rsTemp!CertSN
        ObjItem.SubItems(2) = "" & rsTemp!CertDN
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If lvwCert.ListItems.Count > 0 Then
        lvwCert.ListItems(1).Selected = True
        '填充数字签名启停用记录
        Call FillLogOnOff(Val(Mid(lvwCert.SelectedItem.Key, 2)))
    Else
        Call FillLogOnOff(0)
    End If
    
    '显示指定人员的性质
    gstrSQL = "select A.人员性质,B.说明 from 人员性质说明 A,人员性质分类 B where A.人员性质=B.名称 and A.人员ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    
    Do Until rsTemp.EOF
        lvw人员性质_S.ListItems.Add , "C" & rsTemp("人员性质"), rsTemp("人员性质")
        lvw人员性质_S.ListItems("C" & rsTemp("人员性质")).SubItems(1) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '显示照片
    strTempFile = Sys.ReadLobV2("人员照片", "照片", "人员ID=[1]", "", Val(Mid(lvwMain.SelectedItem.Key, 2)))
    img照片.Picture = LoadPicture(strTempFile)
    img照片.ToolTipText = GetPictureInfo(img照片.Picture)
    '删除该临时文件
    If img照片.ToolTipText <> "无照片" Then
        Kill strTempFile
    End If
    
    img照片.Left = pic镜框.ScaleLeft
    img照片.Top = pic镜框.ScaleTop
    
    '显示个人简介
    txt说明.Text = lvwMain.SelectedItem.Tag
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMenu()
    Dim blnEnabled As Boolean
    Dim lng撤档时间 As Long
    Dim i As Long
    
    blnEnabled = Not tvwMain_S.SelectedItem Is Nothing
    If blnEnabled = True Then
        blnEnabled = tvwMain_S.SelectedItem.Key <> "Root"
    End If
    Toolbar1.Buttons("New").Enabled = blnEnabled
    mnuEditNew.Enabled = blnEnabled
    
    blnEnabled = Not (lvwMain.ListItems.Count = 0 Or lvwMain.SelectedItem Is Nothing)
    Toolbar1.Buttons("Modify").Enabled = blnEnabled
    Toolbar1.Buttons("Delete").Enabled = blnEnabled
'    Toolbar1.Buttons("Start").Enabled = blnEnabled
'    Toolbar1.Buttons("Stop").Enabled = blnEnabled
    mnuEditDelete.Enabled = blnEnabled
    mnuEditModify.Enabled = blnEnabled
    mnuEditExtend.Enabled = blnEnabled
    mnuFileReport.Enabled = blnEnabled
    mnuEditAdjust.Enabled = blnEnabled
    mnuEditRole.Enabled = blnEnabled
    mnuEditStart.Enabled = blnEnabled
    mnuEditStop.Enabled = blnEnabled
    
    mnuPlugIn.Enabled = blnEnabled
    mnuShortMenu(4).Enabled = blnEnabled
    Toolbar1.Buttons("plugIn").Enabled = blnEnabled
    
    If Not lvwMain.SelectedItem Is Nothing Then
        For i = 2 To lvwMain.ColumnHeaders.Count
            If lvwMain.ColumnHeaders(i).Text = "撤档时间" Then
                lng撤档时间 = i
                Exit For
            End If
        Next
        If lvwMain.SelectedItem.ListSubItems(lng撤档时间 - 1) <> "3000-01-01" Then
            mnuEditStart.Enabled = True
            Toolbar1.Buttons("Start").Enabled = True
            
            mnuEditDelete.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditAdjust.Enabled = False
            mnuEditRole.Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
        Else
            mnuEditStart.Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            
            mnuEditDelete.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditExtend.Enabled = True
            mnuEditStop.Enabled = True
            mnuEditAdjust.Enabled = True
            mnuEditRole.Enabled = True
            Toolbar1.Buttons("Modify").Enabled = True
            Toolbar1.Buttons("Delete").Enabled = True
            Toolbar1.Buttons("Stop").Enabled = True
        End If
    End If
    
    EnablePrint lvwMain.ListItems.Count <> 0
    
    '数字证书功能
    mnuEditRegCert.Enabled = blnEnabled
    mnuEditImportCertPic.Enabled = Not lvwCert.SelectedItem Is Nothing
    mnuEditViewCert.Enabled = Not lvwCert.SelectedItem Is Nothing
    mnuEditDelCert.Enabled = Not lvwCert.SelectedItem Is Nothing
        
    stbThis.Panels(2).Text = "人员列表共显示有" & lvwMain.ListItems.Count & "个人员。"
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'功能:设置打印和预鉴按钮的有效值
'参数:blnEnabled 有效值
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub Set权限控制()
'功能:1.由于有的用户权限不够,故使一些菜单项或按钮不可见
'     2.电子签名控制和初始化
    Dim rsTmp As New ADODB.Recordset
    Dim lngSys As Long
    
    '获取使用的电子签名认证中心
    On Error GoTo ErrHandle
    mintCA = 0
    gstrSQL = "Select 参数值 From Zlparameters Where 系统 = " & glngSys & " And Nvl(私有, 0) = 0 And 模块 Is Null  And 参数号=25"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mintCA = Val(NVL(rsTmp!参数值))
    End If
    
    mnuShortSign(0).Visible = False
    mnuShortSign(1).Visible = False
    
    '电子签名接口控制
    If mintCA <> 0 And InStr(mstrPrivs, "数字证书注册") > 0 Then
        On Error Resume Next
        Set mobjESign = CreateObject("zl9ESign.clsESign")
        Err.Clear: On Error GoTo 0
        If Not mobjESign Is Nothing Then
            If mobjESign.Initialize(gcnOracle, glngSys) Then
                mnuEdit_1.Visible = True
                mnuEditRegCert.Visible = True
                mnuEditViewCert.Visible = True
                mnuEditImportCertPic.Visible = True
                
                mnuEdit_2.Visible = True
                mnuEditDelCert.Visible = True
                
                '数字证书启停用
                mnuEditSignLine1.Visible = True
                mnuEditSignOn.Visible = True
                mnuEditSignOff.Visible = True
                
                lbl标题(3).Visible = True
                lvwCert.Visible = True
            Else
                Set mobjESign = Nothing
            End If
        End If
    End If
    
    '数字证书启停用
    Toolbar1.Buttons("SignLine").Visible = mnuEditSignOn.Visible
    Toolbar1.Buttons("Sign").Visible = mnuEditSignOn.Visible
    pic证书.Visible = mnuEditSignOn.Visible
    lvwLogOnOff.Visible = mnuEditSignOn.Visible
    
    '权限控制
    If InStr(mstrPrivs, "增删改") = 0 And InStr(mstrPrivs, "数字证书注册") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuEditViewCert.Visible = False
        mnuEditImportCertPic.Visible = False
        
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortMenu(3).Visible = False
        mnuShortsplit1.Visible = False
        
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    ElseIf InStr(mstrPrivs, "增删改") = 0 Then
        mnuEditNew.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEdit_1.Visible = False
        
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortMenu(3).Visible = False
        mnuShortsplit1.Visible = False
        
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    ElseIf InStr(mstrPrivs, "数字证书注册") = 0 Then
        mnuEdit_1.Visible = False
        mnuEditRegCert.Visible = False
        mnuEditViewCert.Visible = False
        mnuEditImportCertPic.Visible = False
        mnuEdit_2.Visible = False
        mnuEditDelCert.Visible = False
    ElseIf InStr(mstrPrivs, "所有部门") = 0 Then
        mnuViewShow.Checked = True
        mnuViewShow.Enabled = False
    End If
    
    If InStr(mstrPrivs, ";扩展信息维护;") = 0 Then
        mnuEditExtend.Visible = False
    End If
    
    gstrSQL = "Select 编号 from zlSystems"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsTmp.Filter = "编号=300"
    If Not rsTmp.EOF Then
'        mnuFileReport.Enabled = True
'        mnuFileReport.Visible = True
        mnuFileFile.Enabled = True
        mnuFileFile.Visible = True
        mnuSplit1.Visible = True
    Else
        '非病案系统，不用显示报表
'        mnuFileReport.Enabled = False
'        mnuFileReport.Visible = False
        mnuFileFile.Enabled = False
        mnuFileFile.Visible = False
'        mnusplit1.Visible = False
    End If
    
    
    If glngSys = 100 Then
        mnuFileReport.Enabled = True
        mnuFileReport.Visible = True
    Else
        mnuFileReport.Enabled = False
        mnuFileReport.Visible = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set照片缩放()
'功能：设置人员照片是否随着镜框自动变化
    Dim bln缩放 As Boolean
    
    bln缩放 = mnuViewStretch.Checked
    img照片.Stretch = bln缩放
    
    img照片.Left = pic镜框.ScaleLeft
    img照片.Top = pic镜框.ScaleTop
    
    If bln缩放 = True Then
        '不需要调整位置
        img照片.MousePointer = vbArrow
        img照片.Width = pic镜框.ScaleWidth
        img照片.Height = pic镜框.ScaleHeight
    Else
        '需要调整位置
        img照片.MousePointer = vbSizeAll
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function FindKey(ByVal strKey As String) As Boolean
    Dim nodTmp As Node
    For Each nodTmp In tvwMain_S.Nodes
        If nodTmp.Key = strKey Then
            FindKey = True
            Exit Function
        End If
    Next
End Function

Private Function SaveSignPIC(ByVal lng人员id As Long, ByVal strFileName As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle
    blnOk = Sys.SaveLob(100, 15, lng人员id, strFileName)
    SaveSignPIC = blnOk
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckIsUser(ByVal lngUserID As Long) As Boolean
'检查当前人员有无对应用户名
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    strTmp = "select count(人员id) rec from 上机人员表 where 人员id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Caption, lngUserID)
    If rsTmp!Rec = 1 Then
        CheckIsUser = True
    ElseIf rsTmp!Rec > 1 Then
        MsgBox "该人员存在多个登录账户，上级人员表数据存在问题，请联系管理员进行处理！", vbInformation, gstrSysName
    Else
        MsgBox "该人员尚未创建用户，请到管理工具中创建该人员的登录用户！", vbInformation, gstrSysName
    End If
    rsTmp.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
            gstrSQL = "Select a.Id, a.姓名, b.部门id" & _
                       " From 人员表 A, 部门人员 B " & _
                       " Where a.Id = b.人员id And b.缺省 = 1 and (a.编号 like [1] or a.姓名 like [1] or a.简码 like [1]) "
            
            If mnuViewShowStop.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "人员查询", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " 没有找到符合条件的信息！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " 已经定位完所有找到的信息，请重新输入条件！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    With frmPresManage.tvwMain_S
        .Nodes("C" & mrsFind("部门ID")).Selected = True
        .SelectedItem.EnsureVisible
        frmPresManage.FillList "C" & mrsFind("部门ID")
    End With
        
    With frmPresManage.lvwMain
        .ListItems("C" & mrsFind("ID")).Selected = True
        .SelectedItem.EnsureVisible
        frmPresManage.lvwMain_ItemClick .SelectedItem
    End With
End Sub

Private Sub FillLogOnOff(ByVal lngId As Long)
'功能：填充数字证书启停用记录
'参数：
'  lngID：证书ID

    Dim ObjItem As ListItem
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    'XMLType字段读取
    If lngId <> 0 Then
        strSQL = "Select b.Stop_Time, b.Start_Time " & vbCr & _
                 "From 人员证书记录 A, " & vbCr & _
                 "     Xmltable('/root/records' Passing a.停用记录 Columns Stop_Time Varchar2(30) Path '/records/stop_time'," & vbCr & _
                 "              Start_Time Varchar2(30) Path '/records/start_time') B " & vbCr & _
                 "Where a.Id = [1] And Nvl(a.是否停用,0) =1 " & vbCr & _
                 "Order By To_Date(b.Stop_Time, 'yyyy-mm-dd hh24:mi:ss') Desc "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取人员证书记录的启停用记录", lngId)
    End If
    With Me.lvwLogOnOff
        .ListItems.Clear
        i = 1
        If Not rsTmp Is Nothing Then
            Do While rsTmp.EOF = False
                strTmp = NVL(rsTmp!Stop_Time)
                Set ObjItem = .ListItems.Add(, "_" & i, Format(strTmp, "yyyy-mm-dd hh:MM:ss"))
                strTmp = NVL(rsTmp!Start_Time, "1")
                If CDate(strTmp) >= CDate("1990-01-01 00:00:00") Then
                    ObjItem.SubItems(1) = Format(strTmp, "yyyy-mm-dd hh:MM:ss")
                End If
                i = i + 1
                rsTmp.MoveNext
            Loop
        End If
        If .ListItems.Count > 0 Then .ListItems(1).Selected = True
    End With
    
    '处理菜单状态
    Toolbar1.Buttons("Sign").Enabled = lvwCert.ListItems.Count > 0 And mblnCAOnOff
    If lvwCert.ListItems.Count <= 0 Then
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = False
        Exit Sub
    End If
    If mblnCAOnOff = False Or lvwCert.SelectedItem.Index <= 0 Or lvwCert.SelectedItem.Index > 1 Then
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = False
        Exit Sub
    End If
    If lvwLogOnOff.ListItems.Count <= 0 Then
        '默认为启用状态，只能停用操作
        mnuEditSignOn.Enabled = False
        mnuEditSignOff.Enabled = True
    Else
        strTmp = (lvwLogOnOff.ListItems(1).SubItems(1))
        mnuEditSignOn.Enabled = strTmp = ""
        mnuEditSignOff.Enabled = Not mnuEditSignOn.Enabled
    End If
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

