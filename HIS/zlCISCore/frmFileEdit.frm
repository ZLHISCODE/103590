VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmFileEdit 
   Caption         =   "病历文件"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   Icon            =   "frmFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar stbInfo 
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   706
            MinWidth        =   706
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLstItem 
      Left            =   7800
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":08CA
            Key             =   "元素"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFile 
      Height          =   6495
      Left            =   480
      ScaleHeight     =   6435
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   1800
      Width           =   6735
      Begin zl9CISCore.ctrlPatientFile ProFile1 
         Height          =   5175
         Left            =   600
         TabIndex        =   0
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   9128
         AllowEdit       =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ilstbrMain 
      Left            =   2880
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":09DC
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":0BF8
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":0E14
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1030
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":124C
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1468
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1684
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":189E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":2110
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":2330
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":254A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":2764
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":2EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":30F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":3312
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":352C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":3746
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":3EC0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":463A
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":4854
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":4A6E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":50E8
            Key             =   "Add"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   4680
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5302
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5522
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5742
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5962
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5B82
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5DA2
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":5FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":61DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":63FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":661C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":683C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":6A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":6C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":6E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":70AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":7824
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":7A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":7C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":7E72
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":808C
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":8806
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":8F80
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":919A
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":93B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":9A2E
            Key             =   "Add"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   11220
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   660
      Width1          =   9000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   660
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11100
         _ExtentX        =   19579
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
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Description     =   "编辑"
               Object.ToolTipText     =   "保存病历文件"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "打印预览病历"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印病历"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "历史"
               Key             =   "历史"
               Object.ToolTipText     =   "查看病历修订历史"
               Object.Tag             =   "历史"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "编辑"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全文"
               Key             =   "全文"
               Description     =   "编辑"
               Object.ToolTipText     =   "选择病历全文示范模板"
               Object.Tag             =   "全文"
               ImageKey        =   "Sample"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "元素"
               Key             =   "元素"
               Description     =   "编辑"
               Object.ToolTipText     =   "选择元素示范模板"
               Object.Tag             =   "元素"
               ImageKey        =   "History"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "插入"
               Key             =   "插入"
               Description     =   "编辑"
               Object.ToolTipText     =   "在当前元素之前插入新的元素"
               Object.Tag             =   "插入"
               ImageKey        =   "Insert"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Description     =   "编辑"
               Object.ToolTipText     =   "将当前元素从病历中删去"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "添加"
               Key             =   "添加"
               Description     =   "编辑"
               Object.ToolTipText     =   "在病历末尾添加新的内容（如病程记录、护理记录等）"
               Object.Tag             =   "添加"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Description     =   "编辑"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "引入"
               Key             =   "复制"
               Description     =   "编辑"
               Object.ToolTipText     =   "引入最近的病历文本及诊断"
               Object.Tag             =   "引入"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "符号"
               Key             =   "符号"
               Description     =   "编辑"
               Object.ToolTipText     =   "在文本中插入特殊字符"
               Object.Tag             =   "符号"
               ImageKey        =   "SpecChar"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "文本"
               Key             =   "文本"
               Description     =   "编辑"
               Object.ToolTipText     =   "显示所见单文本"
               Object.Tag             =   "文本"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "转储"
               Key             =   "转储"
               Description     =   "编辑"
               Object.ToolTipText     =   "将当前所见单的内容转换成文本"
               Object.Tag             =   "转储"
               ImageKey        =   "toText"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "编辑"
               Key             =   "编辑"
               Description     =   "编辑"
               Object.ToolTipText     =   "编辑病历标记图"
               Object.Tag             =   "编辑"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找病人病历"
               Object.Tag             =   "查找"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
         Begin VB.TextBox txtTmp 
            Height          =   270
            Left            =   -1000
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   345
            Visible         =   0   'False
            Width           =   270
         End
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2715
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLstItem"
      SmallIcons      =   "iLstItem"
      ColHdrIcons     =   "iLstItem"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwDemo 
      Height          =   2715
      Left            =   5880
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLstItem"
      SmallIcons      =   "iLstItem"
      ColHdrIcons     =   "iLstItem"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   8880
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
      TabIndex        =   2
      Top             =   8760
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFileEdit.frx":9C48
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14737
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPrintSet 
         Caption         =   "打印设置(&U)"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "输出到&Excel"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParamSet 
         Caption         =   "参数设置(&M)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_History 
         Caption         =   "查阅历史(&H)"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "引入文本(&C)"
      End
      Begin VB.Menu mnuEdit_Char 
         Caption         =   "特殊字符(&S)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEdit_Text 
         Caption         =   "显示文本(&D)"
      End
      Begin VB.Menu mnuEdit_Exchange 
         Caption         =   "转换文本(&T)"
      End
      Begin VB.Menu mnuEdit_Map 
         Caption         =   "编辑图形(&G)"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "病历(&A)"
      Begin VB.Menu mnuOrder_Add 
         Caption         =   "全文示范(&A)"
         Begin VB.Menu FileList 
            Caption         =   "无示范文件"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuOrder_Demo 
         Caption         =   "元素示范(&E)"
      End
      Begin VB.Menu mnuOrder_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrder_Imp 
         Caption         =   "导入历史病历(&H)"
      End
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrder_Insert 
         Caption         =   "插入元素(&I)"
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "删除元素(&D)"
      End
      Begin VB.Menu mnuOrder_Rec 
         Caption         =   "添加记录(&R)"
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
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "查找病人(&F)"
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "病人信息(&I)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewDiag 
         Caption         =   "疾病参考(&V)"
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "疾病筛查(&D)"
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
Attribute VB_Name = "frmFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public strPrivs As String       '用户具有本程序的具体权限

Private FileID As String
Private PatientID As String '病人ID
Private CheckID As String '病案ID或挂号单ID
Private PatientType As Integer '0=门诊病人 1=住院病人
Private FileTypeID As String '病历模板文件ID
Private bSample As Boolean '是否示范
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1
Private FileType As Integer '病历种类
Private AdviceID As Long '相关医嘱ID
Private blnAllowEdit As Boolean

Private iCurrElementIndex As Integer '当前元素顺序号
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Function ShowMe(ByVal sFileID As String, sPatientID As String, sCheckID As String, _
    iPatientType As Integer, sFileTypeID As String, bSampleFile As Boolean, frmParent As Object, Optional ByVal bAllowEdit As Boolean = True, Optional iFileType As Integer = 0, _
    Optional ByVal btModal As Byte = 0, Optional ByVal lngAdviceID As Long = 0) As Long
    Dim rsTmp As New ADODB.Recordset, i As Integer
    
    On Error Resume Next
    FileID = sFileID: PatientID = sPatientID: CheckID = sCheckID
    PatientType = iPatientType: FileTypeID = sFileTypeID: bSample = bSampleFile: AdviceID = lngAdviceID
    Me.Tag = FileID  '存放该窗口编辑的病历记录ID
    
    iCurrElementIndex = 1

    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "编码", "编码", 900
        .Add , "类型", "类型", 900
    End With
    With Me.lvwItem
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    
    '获取可选元素列表
    GetElementList
    mnuOrder_Add_FileList
    '获取病人信息
    FileType = 0
    If bSample Then
        Me.Caption = "全文示范"
        stbInfo.Visible = False
    Else
        Me.Caption = "病历文件"
        stbInfo.Visible = True
        If Len(FileID) > 0 Then
            zlDatabase.OpenRecordset rsTmp, "Select 病历名称,病历种类 From 病人病历记录 Where ID=" & FileID, Me.Name
        Else
            zlDatabase.OpenRecordset rsTmp, "Select 名称,种类 From 病历文件目录 Where ID=" & FileTypeID, Me.Name
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0): FileType = rsTmp(1)
        
        zlDatabase.OpenRecordset rsTmp, "Select Nvl(门诊号,0),Nvl(住院号,0),姓名,Nvl(性别,' '),Nvl(年龄,' '),nvl(b.名称,' ') As 科室,nvl(c.名称,' ') As 病区,当前床号 From 病人信息 a,部门表 b,部门表 c Where 病人ID=" & PatientID & " And a.当前科室ID=b.ID(+) And a.当前病区ID=c.ID(+)", "zlCISCore"
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "无病人信息"
        Else
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "门诊号：" & rsTmp(0), "住院号：" & rsTmp(1))
                .Item(1).Text = "姓名：" & rsTmp(2) & "，性别：" & rsTmp(3) & "，年龄：" & rsTmp(4)
                If PatientType = 0 Then
                    .Item(2).Visible = False: .Item(3).Visible = False
                Else
                    .Item(2).Text = "科室：" & rsTmp(5)
                    .Item(3).Text = "病区：" & rsTmp(6) & "，床号：" & NVL(rsTmp(7))
                End If
            End With
            
            Me.Caption = rsTmp(2) + "-" + Me.Caption
        End If
    End If
    
    ProFile1.AllowEdit = bAllowEdit: blnAllowEdit = bAllowEdit
    '处理菜单及工具栏
    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.mnuEdit.Visible = bAllowEdit: Me.mnuOrder.Visible = bAllowEdit
    Me.mnuOrder_1.Visible = bAllowEdit
    For i = 1 To Me.tbrMain.Buttons.Count
        If Me.tbrMain.Buttons(i).Description = "编辑" Then Me.tbrMain.Buttons(i).Visible = bAllowEdit
    Next
    Select Case iFileType
        Case 4 '诊断文书
            Me.mnuOrder_Insert.Visible = False
            Me.mnuOrder_Delete.Visible = False
            Me.mnuOrder_Rec.Visible = False
            Me.mnuOrder_2.Visible = False
            Me.tbrMain.Buttons("插入").Visible = False
            Me.tbrMain.Buttons("删除").Visible = False
            Me.tbrMain.Buttons("添加").Visible = False
    End Select
    
    If bSample Then
        Me.mnuFile_History.Visible = False
        Me.mnuFile_4.Visible = False
        Me.tbrMain.Buttons("历史").Visible = False
    Else
        If Len(FileID) = 0 Then
            Me.mnuFile_History.Enabled = False
        Else
            Me.mnuFile_History.Enabled = True
        End If
        Me.tbrMain.Buttons("历史").Enabled = Me.mnuFile_History.Enabled
    End If

    Set ParentForm = frmParent
    If frmParent Is Nothing Then
        Me.Show IIf(bSample, 1, btModal)
    Else
        Me.Show IIf(bSample, 1, btModal), frmParent
    End If
    ShowMe = CLng(Val(FileID))
End Function

Private Sub FileList_Click(Index As Integer)
    If MsgBox("加载病历示范后，当前病历内容将被覆盖！是否继续？", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在加载病历："
    ProFile1.LoadSample CLng(FileList(Index).Tag), Me.prbRefresh
    ProFile1.SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub Form_Activate()
    If ProFile1.Tag = "" Then Exit Sub
    On Error Resume Next
    
    ProFile1.Tag = ""
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在加载病历："
    ProFile1.ShowFile FileID, PatientID, CheckID, PatientType, FileTypeID, bSample, , Me.prbRefresh, AdviceID
    ProFile1.SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItem.Visible Then Me.lvwItem.Visible = False
    If Me.lvwDemo.Visible Then Me.lvwDemo.Visible = False
    
    ProFile1.SetActiveElement iCurrElementIndex
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    ProFile1.Tag = "Loading"
    '---------权限控制-------------
    'strPrivs = gstrPrivs
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.cbrMain.Visible, Me.cbrMain.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    With stbInfo
        .Left = 0: .Top = Me.cbrMain.Top + lngTools
        .Width = Me.ScaleWidth
        
        If PatientType = 0 Then
            .Panels(1).MINWIDTH = .Width - .Panels(4).MINWIDTH
        Else
            .Panels(1).MINWIDTH = (.Width - .Panels(4).MINWIDTH) / 3
            .Panels(2).MINWIDTH = (.Width - .Panels(4).MINWIDTH) / 3
            .Panels(3).MINWIDTH = (.Width - .Panels(4).MINWIDTH) / 3
        End If
    End With
    With picFile
        .Left = 0: .Top = stbInfo.Top + IIf(Not bSample, stbInfo.Height, 0)
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ProFile1.Modified And ProFile1.AllowEdit Then
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        If MsgBox("是否保存编辑的病历", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            mnuFileSave_Click
        End If
    End If
'    zlCommFun.OpenIme False
    
    Call SaveWinState(Me, App.ProductName)
    
    On Error Resume Next
    ParentForm.EditFile_UnLoad Me.hwnd  '告诉上级窗口编辑已关闭
    ProFile1.Release
End Sub

Private Sub lvwDemo_DblClick()
    Dim blnReadOnly As Boolean, i As Integer
    If Me.lvwDemo.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Me.lvwDemo.Tag
        Case "历史"
            If lvwDemo.SelectedItem.Text = "最新" Then
                FileID = Mid(lvwDemo.SelectedItem.Key, 2)
                blnReadOnly = Not blnAllowEdit
            Else
                FileID = Mid(lvwDemo.SelectedItem.Key, 2) * -1 '以前的版本的记录ID用负数表示
                blnReadOnly = True
            End If
            ProFile1.AllowEdit = Not blnReadOnly
            Me.MousePointer = vbHourglass
            BeginShowProgress "正在加载病历："
            ProFile1.ShowFile FileID, PatientID, CheckID, PatientType, FileTypeID, bSample, , Me.prbRefresh, AdviceID
            ProFile1.SetActiveElement 1
            Me.prbRefresh.Visible = False
            Me.MousePointer = vbDefault
            Me.stbThis.Panels(2).Text = ""
            
            Me.lvwDemo.Visible = False
        
            '处理菜单及工具栏
            Me.mnuFileSave.Visible = Not blnReadOnly: Me.mnuFileSplit(1).Visible = Not blnReadOnly
            Me.mnuEdit.Visible = Not blnReadOnly: Me.mnuOrder.Visible = Not blnReadOnly
            Me.mnuOrder_1.Visible = Not blnReadOnly
            For i = 1 To Me.tbrMain.Buttons.Count
                If Me.tbrMain.Buttons(i).Description = "编辑" Then Me.tbrMain.Buttons(i).Visible = Not blnReadOnly
            Next
        Case "记录"
            With Me.lvwDemo
                ProFile1.AddRecord Mid(.SelectedItem.Key, 2), iCurrElementIndex
                        
                .Visible = False
            End With
        Case Else
            With Me.lvwDemo
                ProFile1.LoadElementSample iCurrElementIndex, Mid(.SelectedItem.Key, 2)
                        
                .Visible = False
            End With
    
            ProFile1.SetActiveElement iCurrElementIndex
    End Select
End Sub

Private Sub lvwDemo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwDemo.SelectedItem Is Nothing Then Exit Sub
        Call lvwDemo_DblClick
    End Select
End Sub

Private Sub lvwDemo_LostFocus()
    Me.lvwDemo.Visible = False
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItem.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItem.SortOrder = IIf(Me.lvwItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItem.SortKey = ColumnHeader.Index - 1
        Me.lvwItem.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItem_DblClick()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItem
        .Visible = False
        
        Me.MousePointer = vbHourglass
        BeginShowProgress "正在刷新病历："
        ProFile1.InsertElement Mid(.SelectedItem.Key, 2), iCurrElementIndex, Me.prbRefresh
        Me.prbRefresh.Visible = False
        Me.MousePointer = vbDefault
    
        Me.stbThis.Panels(2).Text = ""
    End With
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
        Call lvwItem_DblClick
    End Select
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub mnuEdit_Char_Click()
    frmSpecChar.Show vbModal, Me
'    zlCommFun.OpenIme True
'    If gblnOK Then SendKeys frmSpecChar.mstrChar
    If gblnOK Then ProFile1.InsertString iCurrElementIndex, frmSpecChar.mstrChar
    Unload frmSpecChar
End Sub

Private Sub mnuEdit_Copy_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngContentID As Long

    On Error Resume Next
    strSQL = "Select b.ID,b.标题文本,a.病历名称,a.书写日期 From 病人病历记录 a,病人病历内容 b," + _
        "(Select b.元素编码,Max(b.id) As ID From 病人病历记录 a,病人病历内容 b,病历元素目录 c Where a.ID=b.病历记录ID And b.元素编码=c.编码 And a.病人id=" & PatientID & " And " + _
        IIf(PatientType = 1, "主页id=" & CheckID, "挂号单='" & CheckID & "'") & " And (b.元素类型=0 Or c.部件 Like 'ZL9CISCORE.%DIAG%') Group By b.元素编码) c " + _
        "Where a.ID=b.病历记录ID And b.ID=c.ID"
    strSQL = strSQL + " Union Select b.ID,b.标题文本||'('||c.检验项目||')',a.病历名称,a.书写日期 From 病人病历记录 a,病人病历内容 b," + _
        "(Select b.元素编码,nvl(d.标题,' ') As 检验项目,Max(b.id) As ID From 病人病历记录 a,病人病历内容 b,病历元素目录 c,病人病历所见单 d Where a.ID=b.病历记录ID And b.元素编码=c.编码 And d.病历id=b.id(+) And a.病人id=" & PatientID & " And " + _
        IIf(PatientType = 1, "主页id=" & CheckID, "挂号单='" & CheckID & "'") & " And c.部件 Like 'ZL9CISCORE.%SPECRESULT%' And d.控件号=-2 Group By b.元素编码,nvl(d.标题,' ')) c " + _
        "Where a.ID=b.病历记录ID And b.ID=c.ID Order By 书写日期 Desc"
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "病历文本", True, , "请选择该病人最近的病历文本", , , True, _
        Me.Left + Me.tbrMain.Left + IIf(Me.cbrMain.Visible, Me.tbrMain.Buttons("复制").Left, 0), Me.Top + Me.tbrMain.Top + 300 + IIf(Me.cbrMain.Visible, tbrMain.Buttons("复制").Top + Me.tbrMain.Buttons("复制").Height, 0), 0, , , True)
        
    If Not rsTmp Is Nothing Then
        lngContentID = rsTmp("ID"): rsTmp.Close
        Call zlDatabase.OpenRecordset(rsTmp, "Select a.ID,nvl(b.部件,' ') From 病人病历内容 a,病历元素目录 b Where a.元素编码=b.编码 And a.ID=" & lngContentID, Me.Caption)
        
        If Not rsTmp.EOF Then ProFile1.CopyElement iCurrElementIndex, rsTmp("ID"), rsTmp(1)
    End If
End Sub

Private Sub mnuEdit_Exchange_Click()
    If MsgBox("所见单内容将覆盖其文本段内容，是否继续", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    ProFile1.ChangeToText iCurrElementIndex
    
    If Not Me.mnuEdit_Text.Checked Then
        mnuEdit_Text_Click
    Else
        If Not ProFile1.ShowText(iCurrElementIndex, True) Then Me.mnuEdit_Text.Checked = False: Me.tbrMain.Buttons("文本").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuEdit_Map_Click()
    ProFile1.EditElement iCurrElementIndex
End Sub

Private Sub mnuEdit_Text_Click()
    If ProFile1.ShowText(iCurrElementIndex, Not Me.mnuEdit_Text.Checked) Then Me.mnuEdit_Text.Checked = Not Me.mnuEdit_Text.Checked
    Me.tbrMain.Buttons("文本").Value = IIf(Me.mnuEdit_Text.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile_History_Click()
    tbrMain_ButtonClick tbrMain.Buttons("历史")
End Sub

Private Sub mnuFileSave_Click()
    Call SaveFile
End Sub
Private Function SaveFile() As Boolean
    Dim sTmpFileID As String
    With txtTmp
        .Visible = True: .SetFocus: DoEvents: .Visible = False
    End With
    
    sTmpFileID = ProFile1.SaveFile
    If Len(sTmpFileID) > 0 Then
        FileID = sTmpFileID: Me.Tag = FileID '存放该窗口编辑的病历记录ID
        SaveFile = True
    
        Me.mnuFile_History.Enabled = True
        Me.tbrMain.Buttons("历史").Enabled = True
    Else
        SaveFile = False
    End If
    ProFile1.SetActiveElement iCurrElementIndex
End Function
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

Private Sub mnuOrder_Delete_Click()
    Me.MousePointer = vbHourglass
    Me.prbRefresh.Value = 0: BeginShowProgress "" '"正在刷新病历："
    ProFile1.DeleteElement iCurrElementIndex, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Demo_Click()
    tbrMain_ButtonClick tbrMain.Buttons("元素")
End Sub

Private Sub mnuOrder_Imp_Click()
    Dim lngImpId As Long    '要导入的病历记录ID
    
    '获取病历文件
    lngImpId = GetFileId(CLng(FileTypeID))
    
    If lngImpId = 0 Then Exit Sub
    If MsgBox("导入历史病历后，当前病历内容将被覆盖！是否继续？", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Err = 0: On Error Resume Next
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在加载病历："
    ProFile1.LoadSample lngImpId, Me.prbRefresh, False
    ProFile1.SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Insert_Click()
    tbrMain_ButtonClick tbrMain.Buttons("插入")
End Sub

Private Sub mnuOrder_Rec_Click()
    tbrMain_ButtonClick tbrMain.Buttons("添加")
End Sub

Private Sub mnuPreview_Click()
    Dim frmPreview As frmCasePrint
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Len(FileID) = 0 Then
        If MsgBox("该病历是新增的，打印之前系统将保存该份病历。是否继续", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1.Modified And ProFile1.AllowEdit Then _
            If MsgBox("打印之前是否保存该份病历", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
    End If
    If bSample Then
        Set frmPreview = New frmCasePrint
        PrintOutCase Me, frmPreview, 0, True, 1, 0, FileID, False, 0, 1
        frmPreview.Preview Me, 0, True, 1, 0, FileID, False, 0, 1
    Else
        If 1 * FileID > 0 Then
            Set frmPreview = New frmCasePrint
            PrintOutCase Me, frmPreview, FileType, True, -1 * FileID, CLng(PatientID), CheckID, False, 0, 1
            frmPreview.Preview Me, FileType, True, -1 * FileID, CLng(PatientID), CheckID, False, 0, 1
        Else
            Set frmPreview = New frmCasePrint
            PrintOutCase Me, frmPreview, FileType, True, 1, 0, CLng(FileID), False, 0, 1
            frmPreview.Preview Me, FileType, True, 1, 0, CLng(FileID), False, 0, 1
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Len(FileID) = 0 Then
        If MsgBox("该病历是新增的，打印之前系统将保存该份病历。是否继续", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1.Modified And ProFile1.AllowEdit Then _
            If MsgBox("打印之前是否保存该份病历", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
    End If
'            If MsgBox("准备打印病历，打印机准备就续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intPage = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(Me) Then
        MsgBox "打印机初始化失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.stbThis.Panels(2).Text = "正在向打印机 " & Printer.DeviceName & " 输出..."
    If bSample Then
        PrintOutCase Me, Printer, 0, True, 1, 0, FileID, False, 0, 1
    Else
        If 1 * FileID > 0 Then
            PrintOutCase Me, Printer, FileType, True, -1 * FileID, CLng(PatientID), CheckID, False, 0, 1
        Else
            PrintOutCase Me, Printer, FileType, True, 1, 0, CLng(FileID), False, 0, 1
        End If
    End If
    'WinNT自定义纸张处理
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(Me)
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuPrintSet_Click()
    frmPrintSet.Show vbModal
End Sub

Private Sub mnuRefresh_Click()
    If MsgBox("本操作将重新调入保存的病历，此前所作" + Chr(13) + "的修改如果未保存将被放弃，是否继续？", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在加载病历："
    ProFile1.ShowFile FileID, PatientID, CheckID, PatientType, FileTypeID, bSample, , Me.prbRefresh, AdviceID
    ProFile1.SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Me.stbThis.Panels(2).Text = ""
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

Private Sub mnuViewDiag_Click()
    frmDiagHelp.ShowMe vbModal, Me
End Sub

Private Sub mnuViewDoctor_Click()
    If PatientType = 0 Then
        frmDiagnotor.ShowMe vbModal, Me, CLng(PatientID), False, , CheckID
    Else
        frmDiagnotor.ShowMe vbModal, Me, CLng(PatientID), True, CLng(CheckID)
    End If
End Sub

Private Sub ParentForm_Unload(Cancel As Integer)
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

Private Sub ProFile1_ElementGotFocus(ByVal ElementIndex As Integer, ByVal ElementType As Integer)
    iCurrElementIndex = ElementIndex
    
    ShowEditMenu ElementType
End Sub

Private Sub ProFile1_Resize()
    If Me.Width < ProFile1.Width Then Me.Width = ProFile1.Width
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "预览"
            mnuPreview_Click
        Case "打印"
            mnuPrint_Click
        Case "保存"
            mnuFileSave_Click
        Case "插入"
            With Me.lvwItem
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwDemo.Visible = False
                .SetFocus
            End With
        Case "全文"
            Call PopupButtonMenu(Me.tbrMain, Button, Me.mnuOrder_Add)
        Case "历史"
            With Me.lvwDemo
                GetFileHistory
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "元素"
            With Me.lvwDemo
                GetElementDemoList ProFile1.ElementID(iCurrElementIndex)
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "删除"
            mnuOrder_Delete_Click
        Case "添加"
            With Me.lvwDemo
                GetAddFile
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "复制"
            mnuEdit_Copy_Click
        Case "符号"
            mnuEdit_Char_Click
        Case "文本"
            mnuEdit_Text_Click
        Case "转储"
            mnuEdit_Exchange_Click
        Case "编辑"
            mnuEdit_Map_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
            mnuExit_Click
    End Select
End Sub

Private Sub ShowEditMenu(ElementType As Integer)
    If Not ProFile1.AllowEdit Then Exit Sub
    Select Case ElementType
        Case 2 '所见单
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = True
            Me.tbrMain.Buttons("文本").Value = IIf(ProFile1.IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("转储").Enabled = True
            Me.tbrMain.Buttons("符号").Enabled = False 'True
            Me.tbrMain.Buttons("编辑").Enabled = False
        Case 3 '标记图
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = False
            Me.tbrMain.Buttons("文本").Value = tbrUnpressed
            Me.tbrMain.Buttons("转储").Enabled = False
            Me.tbrMain.Buttons("符号").Enabled = False
            Me.tbrMain.Buttons("编辑").Enabled = True
        Case 4 '专用纸
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = True
            Me.tbrMain.Buttons("文本").Value = IIf(ProFile1.IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("转储").Enabled = True
            Me.tbrMain.Buttons("符号").Enabled = False 'True
            Me.tbrMain.Buttons("编辑").Enabled = False
        Case Else
            Me.tbrMain.Buttons("复制").Enabled = IIf(ElementType = 0, True, False)
            Me.tbrMain.Buttons("文本").Enabled = False
            Me.tbrMain.Buttons("文本").Value = tbrUnpressed
            Me.tbrMain.Buttons("转储").Enabled = False
            Me.tbrMain.Buttons("符号").Enabled = IIf(ElementType = 0 Or ElementType = -5, True, False) 'True
            Me.tbrMain.Buttons("编辑").Enabled = False
    End Select
    
    Me.mnuEdit_Copy.Enabled = Me.tbrMain.Buttons("复制").Enabled
    Me.mnuEdit_Char.Enabled = Me.tbrMain.Buttons("符号").Enabled
    Me.mnuEdit_Map.Enabled = Me.tbrMain.Buttons("编辑").Enabled
    Me.mnuEdit_Text.Enabled = Me.tbrMain.Buttons("文本").Enabled
    Me.mnuEdit_Text.Checked = IIf(Me.tbrMain.Buttons("文本").Value = tbrPressed, True, False)
    Me.mnuEdit_Exchange.Enabled = Me.tbrMain.Buttons("转储").Enabled
    
    Me.mnuViewDoctor.Visible = Not bSample
    
    If bSample Then
        Me.mnuEdit_Copy.Visible = False
        Me.tbrMain.Buttons("复制").Visible = False
    End If
End Sub

Private Sub GetElementList()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    Me.lvwItem.ListItems.Clear
    Err = 0: On Error GoTo errHand
    Select Case PatientType
        Case 0
            gstrSql = "select I.ID,I.编码,I.名称,I.类型 from 病历元素目录 I where substr(I.适用,1,1)='1' And (类型>=0 Or 类型=-5) order by I.编码"
        Case 1
            gstrSql = "select I.ID,I.编码,I.名称,I.类型 from 病历元素目录 I where substr(I.适用,2,1)='1' And (类型>=0 Or 类型=-5) order by I.编码"
        Case 2
            gstrSql = "select I.ID,I.编码,I.名称,I.类型 from 病历元素目录 I where substr(I.适用,3,1)='1' And (类型>=0 Or 类型=-5) order by I.编码"
        Case 3
            gstrSql = "select I.ID,I.编码,I.名称,I.类型 from 病历元素目录 I where substr(I.适用,4,1)='1' And (类型>=0 Or 类型=-5) order by I.编码"
    End Select
    With rsTemp
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        If .BOF Or .EOF Then
            MsgBox "未建立用于诊疗单据的病历元素！", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "元素": objItem.SmallIcon = "元素"
            objItem.SubItems(Me.lvwItem.ColumnHeaders("编码").Index - 1) = !编码
            strTemp = Switch(!类型 = 0, "文本段", !类型 = 1, "附加表", !类型 = 2, "所见单", !类型 = 3, "标记图", !类型 = 4, "专用纸", _
                            !类型 = -1, "书写签名", !类型 = -2, "当前日期", !类型 = -3, "当前时间", !类型 = -4, "段落标题", !类型 = -5, "普通文本")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("类型").Index - 1) = strTemp
            .MoveNext
        Loop
        Me.lvwItem.ListItems(1).Selected = True
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuOrder_Add_FileList()
    Dim rsFileList As New ADODB.Recordset
    Dim i As Integer, iNum As Integer
    
    On Error Resume Next
    '清除文件清单
    iNum = FileList.Count
    FileList(0).Visible = True
    For i = 1 To iNum - 1
        Unload FileList(i)
    Next
    
    If Len(FileTypeID) = 0 Then
        If bSample Then
            zlDatabase.OpenRecordset rsFileList, "Select 文件ID From" + _
            " 病历示范目录 Where ID=" & FileID, Me.Caption
            
            FileTypeID = rsFileList(0)
        Else
            zlDatabase.OpenRecordset rsFileList, "Select 文件ID From" + _
            " 病人病历记录 Where ID=" & FileID, Me.Caption
            
            FileTypeID = rsFileList(0)
        End If
    End If
    
    zlDatabase.OpenRecordset rsFileList, "Select a.ID,a.名称 From 病历示范目录 a" + _
        " Where a.文件ID=" & FileTypeID & " And a.类型=1" + _
        IIf(bSample, " And a.ID<>" & FileID, "") + _
        IIf(bSample, "", " And (a.科室ID=" & UserInfo.部门ID & " Or" + _
        " a.科室ID Is Null)"), Me.Caption
    If rsFileList.EOF Then Exit Sub
    
    i = 1
    Do While Not rsFileList.EOF
        Load FileList(FileList.Count)
        With FileList(FileList.Count - 1)
            .Caption = "&" & i & " " & rsFileList("名称")
            .Tag = rsFileList("ID")
            .Enabled = True
            .Visible = True
        End With
        
        i = i + 1
        rsFileList.MoveNext
    Loop
    
    FileList(0).Visible = False
End Sub

Private Sub GetElementDemoList(ByVal ElementID As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo errHand
    zlDatabase.OpenRecordset rsTemp, "Select a.ID,a.名称,a.说明 From 病历示范目录 a" + _
        " Where a.元素ID=" & ElementID & " And a.类型=2" + _
        IIf(bSample, "", " And (a.科室ID=" & UserInfo.部门ID & " Or" + _
        " a.科室ID Is Null)"), Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = ""
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "说明", "说明", 1800
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "元素": objItem.SmallIcon = "元素"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("说明").Index - 1) = IIf(IsNull(!说明), "", !说明)
            .MoveNext
        Loop
        Me.lvwDemo.Height = (240 + 25) * (.RecordCount + 2)
        Me.lvwDemo.ListItems(1).Selected = True
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'获取病历修订历史
Private Sub GetFileHistory()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    If Len(FileID) = 0 Then Exit Sub
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo errHand
    zlDatabase.OpenRecordset rsTemp, "Select '最新' As 版本,审阅人 As 书写人,审阅日期 As 书写日期,ID From 病人病历记录 Where ID=" & Me.Tag & _
        " Union All Select to_Char(版本序号,'9999') As 版本,书写人,书写日期,ID From 病人病历修订记录 Where 病历记录ID=" & Me.Tag & _
        " Order By 版本 Desc", Me.Caption
    
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = "历史"
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "版本", "版本", 800
        .Add , "书写人", "书写人", 1000
        .Add , "时间", "时间", 1800
    End With
    With Me.lvwDemo
        .ColumnHeaders("版本").Position = 1
        .SortKey = .ColumnHeaders("版本").Index - 1
        .SortOrder = lvwDescending
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !版本)
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("书写人").Index - 1) = IIf(IsNull(!书写人), "", !书写人)
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("时间").Index - 1) = IIf(IsNull(!书写日期), "", !书写日期)
            .MoveNext
        Loop
        Me.lvwDemo.Height = (240 + 25) * (.RecordCount + 2)
        Me.lvwDemo.ListItems(1).Selected = True
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
'获取可附加的病历清单
Private Sub GetAddFile()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo errHand
    zlDatabase.OpenRecordset rsTemp, "Select a.ID,a.名称,a.说明 From 病历文件目录 a" + _
        " Where a.种类=" & FileType & " And Nvl(a.附加,0)=1", Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = "记录"
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "说明", "说明", 1800
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "元素": objItem.SmallIcon = "元素"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("说明").Index - 1) = IIf(IsNull(!说明), "", !说明)
            .MoveNext
        Loop
        Me.lvwDemo.Height = (240 + 25) * (.RecordCount + 2)
        Me.lvwDemo.ListItems(1).Selected = True
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

