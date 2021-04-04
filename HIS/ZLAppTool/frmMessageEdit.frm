VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMessageEdit 
   Caption         =   "消息"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9510
   Icon            =   "frmMessageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "可变化的"
   Begin MSComctlLib.ImageList ilsEdit 
      Left            =   4860
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0442
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0554
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0666
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0778
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":088A
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":099C
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0AAE
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0BC0
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0CD2
            Key             =   "UnderLine"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0DE4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":0EF6
            Key             =   "No"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":124A
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":159E
            Key             =   "Decrease"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":18F2
            Key             =   "Increase"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":1C46
            Key             =   "Fore"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":1F9A
            Key             =   "Back"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   3
      Top             =   2190
      Width           =   4665
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "收件人(&R)"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   1800
      Width           =   1100
   End
   Begin VB.TextBox txtReceive 
      Height          =   300
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1830
      Width           =   4665
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   688
      BandCount       =   2
      VariantHeight   =   0   'False
      _CBWidth        =   9510
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      MinHeight1      =   330
      Width1          =   3000
      Key1            =   "one"
      NewRow1         =   0   'False
      Child2          =   "tlbEdit"
      MinHeight2      =   330
      Width2          =   7200
      Key2            =   "two"
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbOpen 
         Height          =   345
         Left            =   180
         TabIndex        =   13
         Top             =   60
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "Ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "答复"
               Key             =   "Reply"
               Object.ToolTipText     =   "答复"
               Object.Tag             =   "答复"
               ImageKey        =   "Reply"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全部"
               Key             =   "ReplyAll"
               Object.ToolTipText     =   "全部答复"
               Object.Tag             =   "全部"
               ImageKey        =   "ReplyAll"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "转发"
               Key             =   "Forward"
               Object.ToolTipText     =   "转发"
               Object.Tag             =   "转发"
               ImageKey        =   "Forward"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbEdit 
         Height          =   330
         Left            =   3195
         TabIndex        =   9
         Top             =   30
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "剪切"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "复制"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "粘贴"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "splite1"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Font"
               Object.ToolTipText     =   "字体"
               Style           =   4
               Object.Width           =   1600
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Size"
               Object.ToolTipText     =   "字号"
               Style           =   4
               Object.Width           =   1100
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "粗体"
               ImageKey        =   "Bold"
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "斜体"
               ImageKey        =   "Italic"
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnderLine"
               Object.ToolTipText     =   "下划线"
               ImageKey        =   "UnderLine"
               Style           =   1
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fore"
               Object.ToolTipText     =   "字体色"
               ImageKey        =   "Fore"
               Style           =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "背景色"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Object.ToolTipText     =   "靠左对齐"
               ImageKey        =   "Left"
               Style           =   2
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "居中对齐"
               ImageKey        =   "Center"
               Style           =   2
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "靠右对齐"
               ImageKey        =   "Right"
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Sig"
               Object.ToolTipText     =   "项目符号"
               ImageKey        =   "Sig"
               Style           =   1
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Decrease"
               Object.ToolTipText     =   "减少缩进量"
               ImageKey        =   "Decrease"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Increase"
               Object.ToolTipText     =   "增加缩进量"
               ImageKey        =   "Increase"
            EndProperty
         EndProperty
         Begin VB.ComboBox cmbSize 
            Height          =   300
            Left            =   2820
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   90
            Width           =   1000
         End
         Begin VB.ComboBox cmbFont 
            Height          =   300
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   60
            Width           =   1500
         End
      End
      Begin MSComctlLib.Toolbar tlbNew 
         Height          =   345
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "Ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发送"
               Key             =   "Send"
               Object.ToolTipText     =   "发送消息"
               Object.Tag             =   "发送"
               ImageKey        =   "Send"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "Save"
               Object.ToolTipText     =   "保存消息"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重要"
               Key             =   "High"
               Object.ToolTipText     =   "重要性：高"
               Object.Tag             =   "重要"
               ImageKey        =   "High"
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "次要"
               Key             =   "Low"
               Object.ToolTipText     =   "重要性：低"
               Object.Tag             =   "次要"
               ImageKey        =   "Low"
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   3585
      Left            =   240
      TabIndex        =   4
      Top             =   2550
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6324
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MaxLength       =   4000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMessageEdit.frx":22EE
   End
   Begin MSComctlLib.ImageList Ils16 
      Left            =   3690
      Top             =   330
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
            Picture         =   "frmMessageEdit.frx":238B
            Key             =   "Reply"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":24E5
            Key             =   "ReplyAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":263F
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":2799
            Key             =   "High"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":2BEB
            Key             =   "Low"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":303D
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":348F
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":35A1
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":36B3
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   7170
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "主题(&S)："
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   2250
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "收件人："
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   12
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "发件时间："
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "发件人："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   1530
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSend 
         Caption         =   "发送(&E)"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&T)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "清除(&L)"
      End
      Begin VB.Menu mnuEditSel 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "查找(&F)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "查找下一处(&N)"
         Shortcut        =   +{F3}
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
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "格式(&F)"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "字体(&F)"
      End
      Begin VB.Menu mnuFormatFore 
         Caption         =   "字体色(&O)"
      End
      Begin VB.Menu mnuFormatBack 
         Caption         =   "背景色(&B)"
      End
      Begin VB.Menu mnuFormatSig 
         Caption         =   "项目符号(&S)"
      End
      Begin VB.Menu mnuFormatSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatLeft 
         Caption         =   "靠左对齐(&L)"
      End
      Begin VB.Menu mnuFormatCenter 
         Caption         =   "居中对齐(&C)"
      End
      Begin VB.Menu mnuFormatRight 
         Caption         =   "靠右对齐(&R)"
      End
      Begin VB.Menu mnuFormatSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatDecrease 
         Caption         =   "减少缩进量(&D)"
      End
      Begin VB.Menu mnuFormatIncrease 
         Caption         =   "增加缩进量(&I)"
      End
   End
   Begin VB.Menu mnuAct 
      Caption         =   "动作(&A)"
      Begin VB.Menu mnuActReply 
         Caption         =   "答复(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuActReplyAll 
         Caption         =   "全部答复(&L)"
      End
      Begin VB.Menu mnuActForward 
         Caption         =   "转发(&W)"
      End
      Begin VB.Menu mnuActSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActHigh 
         Caption         =   "重要性高(&H)"
      End
      Begin VB.Menu mnuActLow 
         Caption         =   "重要性低(&O)"
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
End
Attribute VB_Name = "frmMessageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'保存当前所编编辑的消息ID
Public mstrID As String     'strID 消息ID。如果为空，表示是新消息
Public mstrByID As String   'strByID 要进行转发、答复的原始ID
Public mlng类型 As Long     '当前用户对应的消息类型
Public mlngMode As Long     '打开方式。1-答复；2-全部答复；3-转发；0-新建邮件

Dim mstr会话ID As String
Dim mblnSend As Boolean     '准备发送
Dim mblnDelete As Boolean   '是否已经删除

Private mrsUser As ADODB.Recordset '存放收件人姓名,用户名,人员性质的记录集

Dim mlngFore As Long          '当前的前景色

Dim mblnChange As Boolean
'保存用于查找的
Public mblnCase As Boolean
Public mblnBegin As Boolean
Public mstrFind As String

Private Const EM_GETLINECOUNT = &HBA

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - 60
    
    '调整第一行各控件的位置
    lbl(0).Top = sngTop + 120
    lbl(0).Left = 60
    lbl(1).Top = lbl(0).Top
    lbl(1).Left = ScaleWidth - lbl(1).Width - 60
    '调整第二行各控件的位置
    cmdReceive.Left = lbl(0).Left
    If mblnSend = False Then
        cmdReceive.Top = lbl(0).Top + lbl(0).Height + 60
    Else
        cmdReceive.Top = sngTop + 120
    End If
    
    txtReceive.Top = cmdReceive.Top + 25
    txtReceive.Left = cmdReceive.Left + cmdReceive.Width + 60
    txtReceive.Width = ScaleWidth - txtReceive.Left - 60
    
    lbl(2).Left = lbl(0).Left
    '调整第三行各控件的位置
    txtSubject.Top = txtReceive.Top + txtReceive.Height + 60
    txtSubject.Left = txtReceive.Left
    txtSubject.Width = txtReceive.Width
    
    lbl(3).Left = lbl(0).Left
    If mblnSend = True Then
        lbl(2).Top = txtReceive.Top + 60
        lbl(3).Top = txtSubject.Top + 60
    Else
        lbl(2).Top = txtReceive.Top
        lbl(3).Top = txtSubject.Top
    End If
        
    '调整编辑控件的位置
    rtfContent.Left = lbl(0).Left
    rtfContent.Top = txtSubject.Top + txtSubject.Height + 60
    rtfContent.Width = ScaleWidth - rtfContent.Left - 60
    rtfContent.Height = sngBottom - rtfContent.Top
    Me.Refresh
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call LoadMessage
    Call SetState
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    SaveWinState Me, App.ProductName
    
    mstr会话ID = ""
    mblnSend = False
    mblnDelete = False

    mblnChange = False
    Set mrsUser = Nothing
End Sub

Private Sub cmbFont_Click()
    rtfContent.SelFontName = cmbFont.Text
End Sub

Private Sub cmbFont_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        rtfContent.SelFontName = cmbFont.Text
        cmbFont.Text = rtfContent.SelFontName
    End If
End Sub

Private Sub cmbSize_Click()
    rtfContent.SelFontSize = cmbSize.Text
End Sub

Private Sub cmbSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(cmbSize.Text) > 200 Then
            rtfContent.SelFontSize = 200
        Else
            rtfContent.SelFontSize = Val(cmbSize.Text)
        End If
        cmbSize.Text = rtfContent.SelFontSize
    End If
End Sub

Private Sub cmdReceive_Click()
    Dim str收件人  As String, strAddressee As String, strLastOne As String
    Dim rsUser As ADODB.Recordset
    Dim strData() As String
    Dim i As Long
    
    Set rsUser = mrsUser
    str收件人 = txtReceive.Text
    If frmSelectReceiver.Get收件人(str收件人, rsUser) = True Then
        '对"str收件人"进行裁剪，使其满足len(str收件人)<=txtReceive.maxlength的条件
        If Len(str收件人) > txtReceive.MaxLength Then
            strAddressee = Mid(str收件人, 1, txtReceive.MaxLength - 3)
            For i = 1 To Len(strAddressee)
                strLastOne = Mid(strAddressee, Len(strAddressee), 1)
                If strLastOne = "," Or strLastOne = ";" Then
                    strAddressee = strAddressee & "..."
                    Exit For
                Else
                    strAddressee = Mid(strAddressee, 1, Len(strAddressee) - 1)
                End If
            Next
        Else
            strAddressee = str收件人
        End If
        
        Set mrsUser = rsUser
        
        txtReceive.Text = strAddressee
        txtSubject.SetFocus
    End If
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub CoolBar1_Resize()
    cmbFont.Left = tlbEdit.Buttons("Font").Left
    cmbFont.Top = tlbEdit.Buttons("Font").Top + 15
    cmbFont.Width = tlbEdit.Buttons("Font").Width
    
    cmbSize.Left = tlbEdit.Buttons("Size").Left
    cmbSize.Top = tlbEdit.Buttons("Size").Top + 15
    cmbSize.Width = tlbEdit.Buttons("Size").Width
    
End Sub

Private Sub mnuActReply_Click()
'答复
    frmMessageEdit.OpenWindow "", mstrID, mlng类型, 1
    Unload Me
End Sub

Private Sub mnuActReplyAll_Click()
'全部答复
    frmMessageEdit.OpenWindow "", mstrID, mlng类型, 2
    Unload Me
End Sub

Private Sub mnuActForward_Click()
'转发消息
    frmMessageEdit.OpenWindow "", mstrID, mlng类型, 3
    Unload Me
End Sub

Private Sub mnuActHigh_Click()
'设置消息的重要性：高
    mnuActHigh.Checked = Not mnuActHigh.Checked
    mnuActLow.Checked = False
    tlbNew.Buttons("High").Value = IIf(mnuActHigh.Checked, tbrPressed, tbrUnpressed)
    tlbNew.Buttons("Low").Value = IIf(mnuActLow.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuActLow_Click()
'设置消息的重要性：低
    mnuActHigh.Checked = False
    mnuActLow.Checked = Not mnuActLow.Checked
    tlbNew.Buttons("High").Value = IIf(mnuActHigh.Checked, tbrPressed, tbrUnpressed)
    tlbNew.Buttons("Low").Value = IIf(mnuActLow.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuEditCut_Click()
'进行剪切
    If rtfContent.SelLength = 0 Then Exit Sub
    
    Clipboard.SetText rtfContent.SelRTF, vbCFRTF
    rtfContent.SelText = ""
End Sub

Private Sub mnuEditCopy_Click()
'复制文本
    If rtfContent.SelLength = 0 Then Exit Sub
    Clipboard.SetText rtfContent.SelRTF, vbCFRTF
End Sub

Private Sub mnuEditPaste_Click()
'粘贴文本
    If Clipboard.GetFormat(vbCFText) = True Then
        If Clipboard.GetText(vbCFRTF) <> "" Then
            rtfContent.SelRTF = Clipboard.GetText(vbCFRTF)
        Else
            rtfContent.SelRTF = Clipboard.GetText(vbCFText)
        End If
    End If
End Sub

Private Sub mnuEditClear_Click()
'清除所选的文本
    rtfContent.SelText = ""
End Sub

Private Sub mnuEditSel_Click()
'全部选择
    rtfContent.SelStart = 0
    rtfContent.SelLength = Len(rtfContent.Text)
End Sub

Private Sub mnuEditFind_Click()
'查找第一个
    Set frmMessageFind.frmMain = Me
    frmMessageFind.Show vbModal, Me
End Sub

Private Sub mnuEditFindNext_Click()
'继续查找
    Call FindText
End Sub

Private Sub mnuFileSave_Click()
'保存消息
    Dim lst As ListItem
    
    '更新主界面
    If SaveMessage(False) = True Then
        '首先判断是否合适的显示位置
        With frmMessageManager
            On Error Resume Next
            If mstrByID <> "" Then
                '对于有转发的，更换原有邮件的图标
                Set lst = .lvwMain.ListItems("C" & mlng类型 & mstrByID)
                If Err <> 0 Then
                    Err.Clear
                Else
                    If lst.Icon <> "Script" Then
                        lst.Icon = "ReadReply"
                        lst.SmallIcon = "ReadReply"
                    End If
                End If
            End If
            If .mlngIndex = 1 Or .mlngIndex = 2 Then Exit Sub
            If mblnDelete = True Then
                '位于草稿
                If .mlngIndex = 0 Then Exit Sub
            Else
                '位于已删除
                If .mlngIndex = 3 Then Exit Sub
            End If
            Set lst = .lvwMain.ListItems.Add(, "C0" & mstrID, txtSubject.Text, "Script", "Script")
            If Err <> 0 Then
                Set lst = .lvwMain.ListItems("C0" & mstrID)
                Err.Clear
            End If
            On Error GoTo 0
            If mnuActHigh.Checked = True Or mnuActLow.Checked = True Then
                lst.SubItems(1) = IIf(mnuActHigh.Checked = True, "高", "低")
                lst.ListSubItems(1).ReportIcon = IIf(mnuActHigh.Checked = True, "High", "Low")
            Else
                lst.SubItems(1) = ""
                lst.ListSubItems(1).ReportIcon = 0
            End If
            lst.Text = txtSubject.Text
            lst.SubItems(2) = gstrUserName
            lst.SubItems(3) = txtReceive.Text
            lst.SubItems(4) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            lst.Tag = 0
            If .lvwMain.ListItems.Count = 1 Then
                lst.Selected = True
                .rtfContent.TextRTF = rtfContent.TextRTF
                .rtfContent.BackColor = rtfContent.BackColor
            End If
            .SetMenu
          End With
    End If
End Sub

Private Sub mnuFileSend_Click()
'发送消息
    Dim lst As ListItem
    
    If mnuFileSend.Enabled = False Then Exit Sub
    
    If SaveMessage(True) = True Then
        '更新主界面
        
        '首先判断是否合适的显示位置
        With frmMessageManager
            On Error Resume Next
            If mstrByID <> "" Then
                '更换原有邮件的图标
                Set lst = .lvwMain.ListItems("C" & mlng类型 & mstrByID)
                If Err <> 0 Then
                    Err.Clear
                Else
                    If lst.Icon <> "Script" Then
                        lst.Icon = "ReadReply"
                        lst.SmallIcon = "ReadReply"
                    End If
                End If
            End If
            
            If .mlngIndex = 1 Then
                Unload Me
                Exit Sub
            End If
            If .mlngIndex = 0 Or .mlngIndex = 3 Then
                '删除草稿消息
                Set lst = .lvwMain.ListItems("C0" & mstrID)
                If Err <> 0 Then
                    Err.Clear
                Else
                    .lvwMain.ListItems.Remove lst.Index
                    If .lvwMain.ListItems.Count > 0 And .lvwMain.SelectedItem Is Nothing Then
                        .lvwMain.ListItems(1).Selected = True
                        .FillText
                        .SetMenu
                    End If
                End If
                Unload Me
                Exit Sub
            End If
            '创建已发送消息
            Set lst = .lvwMain.ListItems.Add(, "C1" & mstrID, txtSubject.Text, "Read", "Read")
            If mnuActHigh.Checked = True Or mnuActLow.Checked = True Then
                lst.ListSubItems(1).ReportIcon = IIf(mnuActHigh.Checked = True, "High", "Low")
            End If
            lst.SubItems(2) = gstrUserName
            lst.SubItems(3) = txtReceive.Text
            lst.SubItems(4) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            lst.Tag = 1
            If .lvwMain.SelectedItem Is Nothing Then
                lst.Selected = True
                .rtfContent.TextRTF = rtfContent.TextRTF
                .rtfContent.BackColor = rtfContent.BackColor
            End If
            .SetMenu
          End With
          Unload Me
          '出现图标
          Call frmMessageRead.UpdateNotify
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
'另存为文件
    On Error Resume Next
    
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

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFormatFont_Click()
'设置字体的各种情况
    On Error Resume Next
    
    With rtfContent
        cdg.FontBold = IIf(IsNull(.SelBold), False, .SelBold)
        cdg.FontItalic = IIf(IsNull(.SelItalic), False, .SelItalic)
        cdg.FontUnderline = IIf(IsNull(.SelUnderline), False, .SelUnderline)
        cdg.FontName = IIf(IsNull(.SelFontName), .Font.Name, .SelFontName)
        cdg.FontSize = IIf(IsNull(.SelFontSize), .Font.Size, .SelFontSize)
        
        cdg.CancelError = True
        cdg.flags = cdlCFScreenFonts
        Err.Clear
        cdg.ShowFont
        
        If Err <> 0 Then
            Err.Clear
        Else
            .SelBold = cdg.FontBold
            .SelItalic = cdg.FontItalic
            .SelUnderline = cdg.FontUnderline
            .SelFontName = cdg.FontName
            .SelFontSize = cdg.FontSize
            
            Call rtfContent_SelChange
        End If
    End With
End Sub

Private Sub mnuFormatFore_Click()
'选择字体颜色

    On Error Resume Next
    cdg.Color = IIf(IsNull(rtfContent.SelColor), mlngFore, rtfContent.SelColor)
    cdg.CancelError = True
    cdg.flags = cdlCCFullOpen Or cdlCCRGBInit
    Err.Clear
    cdg.ShowColor
    
    If Err <> 0 Then
        Err.Clear
    Else
        mlngFore = cdg.Color
        rtfContent.SelColor = cdg.Color
    End If
End Sub

Private Sub mnuFormatBack_Click()
'选择背景颜色
    On Error Resume Next

    cdg.Color = rtfContent.BackColor
    cdg.CancelError = True
    cdg.flags = cdlCCFullOpen Or cdlCCRGBInit
    cdg.ShowColor
    
    If Err <> 0 Then
        Err.Clear
    Else
        rtfContent.BackColor = cdg.Color
    End If
End Sub

Private Sub mnuFormatSig_Click()
'对文本加上项目符号
    mnuFormatSig.Checked = Not mnuFormatSig.Checked
    tlbEdit.Buttons("Sig").Value = IIf(mnuFormatSig.Checked, tbrPressed, tbrUnpressed)
    
    rtfContent.SelBullet = mnuFormatSig.Checked
End Sub

Private Sub mnuFormatLeft_Click()
'文本靠左
    mnuFormatLeft.Checked = True
    mnuFormatRight.Checked = False
    mnuFormatCenter.Checked = False
    tlbEdit.Buttons("Left").Value = tbrPressed
    
    rtfContent.SelAlignment = 0
End Sub

Private Sub mnuFormatCenter_Click()
'文本居中
    mnuFormatLeft.Checked = False
    mnuFormatRight.Checked = False
    mnuFormatCenter.Checked = True
    tlbEdit.Buttons("Center").Value = tbrPressed
    
    rtfContent.SelAlignment = 2
End Sub

Private Sub mnuFormatRight_Click()
'文本居右
    mnuFormatLeft.Checked = False
    mnuFormatRight.Checked = True
    mnuFormatCenter.Checked = False
    tlbEdit.Buttons("Right").Value = tbrPressed
    
    rtfContent.SelAlignment = 1
End Sub

Private Sub mnuFormatIncrease_Click()
' 增加缩进量
    Dim i As Long
    
    i = IIf(rtfContent.SelIndent < rtfContent.Width - 1000, rtfContent.SelIndent, rtfContent.Width - 1000)
    rtfContent.SelIndent = i + 360
End Sub

Private Sub mnuFormatDecrease_Click()
'减少缩进量
    Dim i As Long
    
    i = IIf(rtfContent.SelIndent > 360, rtfContent.SelIndent, 360)
    rtfContent.SelIndent = i - 360
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub rtfContent_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
            Call mnuFileSend_Click
        End If
    End If
End Sub

Private Sub tlbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Cut"
            Call mnuEditCut_Click
        Case "Copy"
            Call mnuEditCopy_Click
        Case "Paste"
            Call mnuEditPaste_Click
        Case "Bold"
            rtfContent.SelBold = (tlbEdit.Buttons("Bold").Value = tbrPressed)
        Case "Italic"
            rtfContent.SelItalic = (tlbEdit.Buttons("Italic").Value = tbrPressed)
        Case "UnderLine"
            rtfContent.SelUnderline = (tlbEdit.Buttons("UnderLine").Value = tbrPressed)
        Case "Fore"
            rtfContent.SelColor = mlngFore
        Case "Back"
            Dim lngColor As Long
            Dim sngLeft As Single, sngTop As Single
            
            '选择颜色
            sngLeft = Me.Left + IIf(CoolBar1.Bands("two").NewRow, 0, IIf(CoolBar1.Bands("two").Position = 1, 0, CoolBar1.Bands("one").Width)) + _
                            Button.Left + Button.Width
            sngTop = Me.Top + CoolBar1.Top + 400 * IIf(CoolBar1.Bands("two").Position = 1, 0, IIf(CoolBar1.Bands("two").NewRow, 1, 0)) _
                            + tlbEdit.Top + Button.Top + Button.Height
            
            lngColor = rtfContent.BackColor
            If frmSelectColor.GetColor(lngColor, Me, sngLeft, sngTop) = True Then
                rtfContent.BackColor = lngColor
            End If
        Case "Left"
            Call mnuFormatLeft_Click
        Case "Center"
            Call mnuFormatCenter_Click
        Case "Right"
            Call mnuFormatRight_Click
        Case "Find"
            Call mnuEditFind_Click
        Case "Sig"
            Call mnuFormatSig_Click
        Case "Increase"
            Call mnuFormatIncrease_Click
        Case "Decrease"
            Call mnuFormatDecrease_Click
    End Select

End Sub

Private Sub tlbEdit_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim lngColor As Long
    
    If Button.Key = "Fore" Then
        '选择颜色
        sngLeft = Me.Left + IIf(CoolBar1.Bands("two").NewRow, 0, IIf(CoolBar1.Bands("two").Position = 1, 0, CoolBar1.Bands("one").Width)) + _
                        Button.Left + Button.Width
        sngTop = Me.Top + CoolBar1.Top + 300 * IIf(CoolBar1.Bands("two").Position = 1, 0, IIf(CoolBar1.Bands("two").NewRow, 1, 0)) _
                        + tlbEdit.Top + Button.Top + Button.Height
        
        lngColor = IIf(IsNull(rtfContent.SelColor), mlngFore, rtfContent.SelColor)
        If frmSelectColor.GetColor(lngColor, Me, sngLeft, sngTop) = True Then
            mlngFore = lngColor
            rtfContent.SelColor = lngColor
        End If
    End If
End Sub

Private Sub tlbNew_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "High"
            Call mnuActHigh_Click
        Case "Low"
            Call mnuActLow_Click
        Case "Help"
            Call mnuhelptopic_Click
        Case "Quit"
            Call mnuFileExit_Click
        Case "Send"
            Call mnuFileSend_Click
        Case "Save"
            Call mnuFileSave_Click
    End Select

End Sub

Private Sub tlbNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tlbOpen_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Reply"
            Call mnuActReply_Click
        Case "ReplyAll"
            Call mnuActReplyAll_Click
        Case "Forward"
            Call mnuActForward_Click
        Case "Help"
            Call mnuhelptopic_Click
        Case "Quit"
            Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub txtReceive_Change()
    '根据内容动态修改文本框的高度
    Dim lngRet As Long
    Dim lngLastLine As Long                   '最后的行数
    Dim lngLineHeight  As Long                        '每行的高度
    
    Set Me.Font = txtReceive.Font
    lngLineHeight = Me.TextHeight("TT")
    lngRet = SendMessage(txtReceive.hWnd, EM_GETLINECOUNT, 0, 0&)
    If lngRet <> lngLastLine Then
        If txtReceive.Height + txtReceive.Top + lngLineHeight > Me.ScaleHeight And lngRet > 1 Then
            If lngLastLine <= lngRet - 1 Then
                Exit Sub '如果已经是最大高度，保持
            End If
            lngLastLine = lngRet - 1    '超过最大高度
        Else
            lngLastLine = lngRet
        End If
        txtReceive.Height = (lngLastLine + 1) * lngLineHeight     '修改高度
    End If
    mblnChange = True
    Call Form_Resize
End Sub

Private Sub txtReceive_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtReceive_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
            Call mnuFileSend_Click
        Else
            txtSubject.SetFocus
        End If
    End If
End Sub

Private Sub txtSubject_Change()
    mblnChange = True
End Sub

Private Sub txtSubject_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtfContent_Change()
    mblnChange = True
End Sub

Private Sub rtfContent_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtfContent_SelChange()
'根据不同的选择内容，更新工具栏上的状态
    With rtfContent
        '字体
        cmbFont.Text = IIf(IsNull(.SelFontName), "", .SelFontName)
        cmbSize.Text = IIf(IsNull(.SelFontSize), "", .SelFontSize)
        tlbEdit.Buttons("Bold").Value = IIf(IIf(IsNull(.SelBold), False, .SelBold), tbrPressed, tbrUnpressed)
        tlbEdit.Buttons("Italic").Value = IIf(IIf(IsNull(.SelItalic), False, .SelItalic), tbrPressed, tbrUnpressed)
        tlbEdit.Buttons("UnderLine").Value = IIf(IIf(IsNull(.SelUnderline), False, .SelUnderline), tbrPressed, tbrUnpressed)
        
        '对齐
        If IsNull(.SelAlignment) Then
            tlbEdit.Buttons("Left").Value = tbrUnpressed
            tlbEdit.Buttons("Center").Value = tbrUnpressed
            tlbEdit.Buttons("Right").Value = tbrUnpressed
        Else
            If .SelAlignment = 0 Then
                tlbEdit.Buttons("Left").Value = tbrPressed
            ElseIf .SelAlignment = 1 Then
                tlbEdit.Buttons("Right").Value = tbrPressed
            Else
                tlbEdit.Buttons("Center").Value = tbrPressed
            End If
        End If
        mnuFormatLeft.Checked = tlbEdit.Buttons("Left").Value
        mnuFormatCenter.Checked = tlbEdit.Buttons("Center").Value
        mnuFormatRight.Checked = tlbEdit.Buttons("Right").Value
        
        '项目符号
        tlbEdit.Buttons("Sig").Value = IIf(IIf(IsNull(.SelBullet), False, .SelBullet), tbrPressed, tbrUnpressed)
        mnuFormatSig.Checked = tlbEdit.Buttons("Sig").Value
    End With
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbNew.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    For Each buttTemp In tlbOpen.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("one").MinHeight = IIf(mblnSend, tlbNew.Height, tlbOpen.Height)
    Form_Resize
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

Public Sub OpenWindow(ByVal strID As String, ByVal strByID As String, Optional ByVal lng类型 As Long, Optional ByVal lngMode As Long)
'功能：根据参数来显示消息编辑窗口
'strID 消息ID。如果为空，表示是新消息
'strByID 要进行转发、答复的原始ID
'lngMode  处理方式。1-答复；2-全部答复；3-转发
    Dim frmMessage As frmMessageEdit
    Dim frmTemp As Form
    
    '查找该消息是否已经打开了编辑窗口
    If strID <> "" Then
        For Each frmTemp In Forms
            If frmTemp.Name = "frmMessageEdit" Then
                If frmTemp.mstrID = strID And frmTemp.mlng类型 = lng类型 Then
                    Set frmMessage = frmTemp
                    Exit For
                End If
            End If
        Next
    End If
    If frmMessage Is Nothing Then
        Set frmMessage = New frmMessageEdit
        frmMessage.mstrID = strID
        frmMessage.mstrByID = strByID
        frmMessage.mlng类型 = lng类型
        frmMessage.mlngMode = lngMode
    End If
    frmMessage.Show , gfrmMain
End Sub

Private Sub SetState()
    '首先设置可见性
    cmdReceive.Visible = mblnSend
    lbl(0).Visible = Not mblnSend
    lbl(1).Visible = Not mblnSend
    lbl(2).Visible = Not mblnSend
    If Not mblnSend Then
        txtReceive.Appearance = 0
        txtReceive.BorderStyle = 0
        txtReceive.BackColor = BackColor
        txtReceive.Enabled = False
        txtSubject.Appearance = 0
        txtSubject.BorderStyle = 0
        txtSubject.BackColor = BackColor
        txtSubject.Enabled = False
    End If
    
    tlbNew.Visible = mblnSend
    tlbOpen.Visible = Not mblnSend
    CoolBar1.Bands("two").Visible = mblnSend
    Set CoolBar1.Bands("one").Child = IIf(mblnSend, tlbNew, tlbOpen)

    '接着设置可用性
    rtfContent.Locked = Not mblnSend
    
    mnuFileSend.Enabled = mblnSend
    mnuFileSave.Enabled = mblnSend
    
    mnuEditCut.Enabled = mblnSend
    mnuEditPaste.Enabled = mblnSend
    mnuEditClear.Enabled = mblnSend
    
    mnuActReply.Enabled = Not mblnSend
    mnuActReplyAll.Enabled = Not mblnSend
    mnuActForward.Enabled = Not mblnSend
    mnuActHigh.Enabled = mblnSend
    mnuActLow.Enabled = mblnSend
    
    mnuFormat.Visible = mblnSend
    
    '装入字体和字号
    Dim lngCount As Long
    
    cmbFont.Clear
    For lngCount = 0 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(lngCount)
    Next
    
    cmbSize.Clear
    cmbSize.AddItem 8
    cmbSize.AddItem 9
    cmbSize.AddItem 10
    cmbSize.AddItem 11
    cmbSize.AddItem 12
    cmbSize.AddItem 14
    cmbSize.AddItem 16
    cmbSize.AddItem 18
    cmbSize.AddItem 20
    cmbSize.AddItem 22
    cmbSize.AddItem 24
    cmbSize.AddItem 26
    cmbSize.AddItem 28
    cmbSize.AddItem 36
    cmbSize.AddItem 48
    cmbSize.AddItem 72
    
    cmbFont.Text = rtfContent.Font.Name
    cmbSize.Text = rtfContent.Font.Size
    
    '设置菜单名称
    mnuEditCut.Caption = "剪切(&T)" & vbTab & "Ctrl+X"
    mnuEditCopy.Caption = "复制(&C)" & vbTab & "Ctrl+C"
    mnuEditPaste.Caption = "粘贴(&P)" & vbTab & "Ctrl+V"
    mnuEditClear.Caption = "清除(&L)" & vbTab & "Del"
    mnuEditSel.Caption = "全选(&A)" & vbTab & "Ctrl+A"
    mnuFileSend.Caption = "发送(&E)" & vbTab & "Ctrl+Enter"
    
    If InStr(frmMessageManager.mstrPrivs, "发送消息") = 0 Then
        mnuFileSend.Visible = False
        mnusplit2.Visible = False
        
        mnuAct.Visible = False
        
        tlbOpen.Buttons("Reply").Visible = False
        tlbOpen.Buttons("ReplyAll").Visible = False
        tlbOpen.Buttons("Forward").Visible = False
        tlbOpen.Buttons("Split").Visible = False
    End If
End Sub

Private Sub LoadMessage()
'装入数据
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset, rsState As ADODB.Recordset
    Dim str发件人 As String
    Dim str收件人 As String
    Dim str姓名  As String
    
    On Error GoTo ErrH
    lngID = Val(IIf(mstrID = "", mstrByID, mstrID)) '
    If lngID = 0 Then
        mblnDelete = False
        mblnSend = True
        Exit Sub '是全新的，用不着装入
    End If
    
    '得到邮件正文
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select A.* ,B.删除,B.状态 from zlmessages A,zlmsgState B " & _
        " where A.ID=B.消息ID and B.消息ID=[1] and b.类型= [2] and B.用户=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, mlng类型, gstrDbUser)
    
    mstr会话ID = rsTemp("会话ID")
    mblnDelete = IIf(rsTemp("删除") = 1, True, False)
    mblnSend = (mstrID = "") Or (mlng类型 <> 2 And mlng类型 <> 1) '新邮件或未发送邮件
    
    lbl(1).Caption = "发送时间：" & IIf(IsNull(rsTemp("时间")), "", Format(rsTemp("时间"), "yyyy-MM-dd HH:mm:ss"))
    str发件人 = IIf(IsNull(rsTemp("发件人")), "", rsTemp("发件人"))
    str收件人 = IIf(IsNull(rsTemp("收件人")), "", rsTemp("收件人"))
    txtSubject.Text = IIf(IsNull(rsTemp("主题")), "", rsTemp("主题"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("内容")), "", rsTemp("内容"))
    rtfContent.BackColor = IIf(IsNull(rsTemp("背景色")), RGB(255, 255, 255), rsTemp("背景色"))
    
    '得到邮递地址
    gstrSQL = "select 类型,用户,身份 from zlmsgstate where 消息ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)

    '创建空的收件人记录集
    Set mrsUser = Rec.CopyNew(Nothing, True, , Array("用户名", adVarWChar, 30, Empty, "姓名", adVarWChar, 30, Empty, "收件人", adVarWChar, 30, Empty))

    '更新界面内容
    Select Case mlngMode
        Case 1 '答复
            rsTemp.Filter = "类型=0 or 类型=1"
            mrsUser.AddNew
            mrsUser.Fields("姓名") = rsTemp("身份")
            mrsUser.Fields("用户名") = rsTemp("用户")
            txtReceive.Text = mrsUser.Fields("姓名")
            If Left(txtSubject.Text, 3) <> "答复：" Then
                txtSubject.Text = "答复：" & txtSubject.Text
            End If
        Case 2 '全部答复
            str姓名 = ""
            If str收件人 = "所有人员" Or str收件人 = "本部门人员" Or str收件人 = "本科室人员" Then
                txtReceive.Text = str收件人
            ElseIf InStr(str收件人, "]") > 0 And InStr(str收件人, "[") > 0 Then
                txtReceive.Text = str收件人
            Else
                If str发件人 <> str收件人 And InStr(str收件人, str发件人 & ",") = 0 And InStr(str收件人, "," & str发件人) = 0 Then
                    txtReceive.Text = str发件人 & "," & str收件人
                Else
                    txtReceive.Text = str收件人
                End If
            End If
            If Left(txtSubject.Text, 3) <> "答复：" Then
                txtSubject.Text = "答复：" & txtSubject.Text
            End If
            rsTemp.Filter = "类型=3 or 类型=2"
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("姓名") = rsTemp("身份")
                mrsUser.Fields("用户名") = rsTemp("用户")
                
                str姓名 = str姓名 & rsTemp("身份") & ","
                rsTemp.MoveNext

            Loop
            If str发件人 <> str姓名 And InStr(str姓名, str发件人 & ",") = 0 And InStr(str姓名, "," & str发件人) = 0 Then
                mrsUser.AddNew
                mrsUser.Fields("姓名") = str发件人
                mrsUser.Fields("用户名") = gstrDbUser
            End If

            
        Case 3 '转发
            txtReceive.Text = ""
            If Left(txtSubject.Text, 3) <> "转发：" Then
                txtSubject.Text = "转发：" & txtSubject.Text
            End If
        Case Else
            txtReceive.Text = str收件人
            rsTemp.Filter = "类型=3 or 类型=2"
            
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("姓名") = rsTemp!身份
                mrsUser.Fields("用户名") = rsTemp!用户
                rsTemp.MoveNext
            Loop
    End Select
    
    
    lbl(0).Caption = "发件人：     " & str发件人
    Me.Caption = "消息  " & IIf(txtSubject.Text = "", "", "-  " & txtSubject.Text)
    If mlngMode <> 0 Then
        '把原件加上区别
        With rtfContent
            .SelStart = 0
            .SelText = vbCrLf & "----------原始消息------------" & vbCrLf
            .SelStart = 2
            .SelLength = Len("----------原始消息------------")
            .SelFontName = "宋体"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelColor = 0
            
            .SelLength = Len(.Text)
            .SelIndent = 720
            
            .SelStart = 0
            .SelLength = 2
            .SelColor = RGB(0, 0, 255)
            .SelFontName = "宋体"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelLength = 0
        End With
    End If
    gstrSQL = "select 状态 from zlmsgstate where 消息ID=" & lngID & " and 类型=" & mlng类型 & " and 用户='" & gstrDbUser & "'"
    Set rsState = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsState, gstrSQL, Me.Caption)
    '修改标记
    gstrSQL = "Zl_Zlmsgstate_Edit(1," & lngID & "," & mlng类型 & ",'" & gstrDbUser & "','" & gstrUserName & "',Null,'1' || substr( '" & rsState!状态 & "',2))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call frmMessageRead.UpdateNotify
    Dim lst As ListItem
    On Error Resume Next
    '改成已读
    Set lst = frmMessageManager.lvwMain.ListItems("C" & mlng类型 & lngID)
    If Not lst Is Nothing Then
        If lst.Icon <> "Script" Then '不是草稿的消息才可能更改
            If InStr(lst.Icon, "Reply") > 0 Then
                lst.Icon = "ReadReply"
                lst.SmallIcon = "ReadReply"
            Else
                lst.Icon = "Read"
                lst.SmallIcon = "Read"
            End If
        End If
        frmMessageManager.SetMenu
    End If
    Err.Clear
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveMessage(ByVal blnSend As Boolean) As Boolean
    Dim rsState As ADODB.Recordset
    Dim lngID As Long
    Dim lng会话ID As Long
    Dim lngCount As Long, i As Long
    Dim strSQL As String
    Dim strData() As String '存储用户名，姓名
    Dim blnTrans As Boolean
    
    '没修改，且只是保存
    
    If blnSend = True And txtReceive.Text = "" And mrsUser Is Nothing Then
        MsgBox "请选择收件人。", vbExclamation, gstrSysName
        cmdReceive.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtReceive.Text, txtReceive.MaxLength, cmdReceive.hWnd, "收件人") = False Then
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtSubject.Text, txtSubject.MaxLength, txtSubject.hWnd, "主题") = False Then
        Exit Function
    End If
    If LenB(StrConv(rtfContent.TextRTF, vbFromUnicode)) > 4000 Then
        MsgBox "正文的字符太多，或者格式太复杂了。", vbExclamation, gstrSysName
        rtfContent.SetFocus
        Exit Function
    End If
    
    If mstrID = "" Then
        lngID = zlDatabase.GetNextId("zlmessages")
        lng会话ID = IIf(mstr会话ID = "", lngID, mstr会话ID)
    Else
        lngID = mstrID
        lng会话ID = mstr会话ID
    End If
    
    ReDim strData(0)
    On Error GoTo errHandle
    '拼接收件人信息
    If Not mrsUser Is Nothing Then
        If mrsUser.State = adStateOpen Then
            If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
            Do Until mrsUser.EOF
                '将收件人用户名，姓名连成字符串格式为：用户名1,姓名1#用户名2,姓名2#用户名3,姓名3
                strSQL = IIf(strSQL = "", "", strSQL & "#") & mrsUser.Fields("用户名") & "," & mrsUser.Fields("姓名")
                '因为传输字符串（gstrSQL）的最大长度是4000，所以如果长度过长，则分段提交
                If zlStr.ActualLen(strSQL) > 3900 Then
                    strData(UBound(strData)) = "Zl_Zlmsgstate_Addaddressee(" & lngID & _
                                                "," & IIf(blnSend, "2", "3") & _
                                                ",'000" & IIf(mnuActHigh.Checked = True, "1'", IIf(mnuActLow.Checked = True, "2'", "0'")) & _
                                                ",'" & strSQL & "')"
                    ReDim Preserve strData(UBound(strData) + 1)
                    strSQL = ""
                End If
                mrsUser.MoveNext
            Loop
        End If
    End If
    strData(UBound(strData)) = "Zl_Zlmsgstate_Addaddressee(" & lngID & _
                                            "," & IIf(blnSend, "2", "3") & _
                                            ",'000" & IIf(mnuActHigh.Checked = True, "1'", IIf(mnuActLow.Checked = True, "2'", "0'")) & _
                                            ",'" & strSQL & "')"
    '1、增加或修改zlmessages中的记录。2、增加或修改发件人的记录。3、删除已有收件人的记录。4、为原件加上答复或转发标志
    gstrSQL = "Zl_Zlmessages_New(" & lngID & _
                                "," & lng会话ID & _
                                ",'" & txtReceive.Text & _
                                "','" & txtSubject.Text & _
                                "','" & Replace(rtfContent.TextRTF, "'", "''") & _
                                "'," & rtfContent.BackColor & _
                                "," & IIf(blnSend, "1", "0") & _
                                ",'" & gstrDbUser & _
                                "','" & gstrUserName & _
                                "','" & IIf(blnSend, "1", "0") & "00" & IIf(mnuActHigh.Checked = True, "1'", IIf(mnuActLow.Checked = True, "2'", "0'")) & _
                                "," & mlngMode & _
                                "," & IIf(mstrByID = "", "NULL", mstrByID) & _
                                "," & mlng类型 & ")"
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '增加收件人的记录
    For i = 0 To UBound(strData)
        Call zlDatabase.ExecuteProcedure(strData(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mstrID = lngID
    mstr会话ID = lng会话ID
    mblnChange = False
    SaveMessage = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub FindText()
    '如果查找内容为空，直接退出
    Dim lngPos As Long, lngStart As Long
    Dim strText As String
    
    If mstrFind = "" Then Exit Sub
    
    strText = rtfContent.Text
    If mblnBegin = False Then
        lngStart = rtfContent.SelStart + rtfContent.SelLength + 1
    Else
        lngStart = 1
    End If
    
    lngPos = InStr(lngStart, IIf(mblnCase = True, strText, UCase(strText)), IIf(mblnCase = True, mstrFind, UCase(mstrFind)))
    
    If lngPos = 0 Then
        MsgBox "查找结束，没找到“" & mstrFind & "”", vbInformation, gstrSysName
    Else
        rtfContent.SelStart = lngPos - 1
        rtfContent.SelLength = Len(mstrFind)
    End If
    
End Sub

Private Sub txtSubject_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
            Call mnuFileSend_Click
        Else
            rtfContent.SetFocus
        End If
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

