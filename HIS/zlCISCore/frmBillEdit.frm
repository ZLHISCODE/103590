VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmBillEdit 
   Caption         =   "诊疗单据"
   ClientHeight    =   9120
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9075
   Icon            =   "frmBillEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6510
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabFile 
      Height          =   350
      Left            =   0
      TabIndex        =   17
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      TabFixedHeight  =   450
      HotTracking     =   -1  'True
      Placement       =   1
      TabMinWidth     =   1764
      ImageList       =   "iLstTab"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbInfo 
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   720
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
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
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLstItem 
      Left            =   8500
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":08CA
            Key             =   "元素"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":09DC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":0F76
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1510
            Key             =   "Template"
         EndProperty
      EndProperty
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
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1AAA
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1CC6
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1EE2
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":20FE
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":231A
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":2536
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":2752
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":2B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":2DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":2FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":31DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":33FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":3618
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":3832
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":3FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":41C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":43E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":45FA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":4814
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":4F8E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":5708
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":5922
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":5B3C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":61B6
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":63D0
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   3840
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":65EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":680A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":6A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":6C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":6E6A
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":708A
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":72AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":74C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":76E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":7904
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":7B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":7D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":7F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":8178
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":8392
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":8B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":8D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":8F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":915A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":9374
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":9AEE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":A268
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":A482
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":A69C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":AD16
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":AF30
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9075
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   645
      Width1          =   9000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   645
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilstbrMain"
         HotImageList    =   "ilstbrMainHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
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
               Object.Visible         =   0   'False
               Description     =   "编辑1"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "全文"
               Key             =   "全文"
               Description     =   "编辑1"
               Object.ToolTipText     =   "选择病历全文示范模板"
               Object.Tag             =   "全文"
               ImageKey        =   "Sample"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "元素"
               Key             =   "元素"
               Description     =   "编辑1"
               Object.ToolTipText     =   "选择元素示范模板"
               Object.Tag             =   "元素"
               ImageKey        =   "History"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "插入"
               Key             =   "插入"
               Description     =   "编辑1"
               Object.ToolTipText     =   "在当前元素之前插入新的元素"
               Object.Tag             =   "插入"
               ImageKey        =   "Insert"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "删除"
               Key             =   "删除"
               Description     =   "编辑1"
               Object.ToolTipText     =   "将当前元素从病历中删去"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Description     =   "编辑"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "引入"
               Key             =   "复制"
               Description     =   "编辑"
               Object.ToolTipText     =   "引入最近的病历文本和诊断"
               Object.Tag             =   "引入"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "符号"
               Key             =   "符号"
               Description     =   "编辑"
               Object.ToolTipText     =   "在文本中插入特殊字符"
               Object.Tag             =   "符号"
               ImageKey        =   "SpecChar"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "文本"
               Key             =   "文本"
               Description     =   "编辑"
               Object.ToolTipText     =   "显示所见单文本"
               Object.Tag             =   "文本"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "转储"
               Key             =   "转储"
               Description     =   "编辑"
               Object.ToolTipText     =   "将当前所见单的内容转换成文本"
               Object.Tag             =   "转储"
               ImageKey        =   "toText"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "编辑"
               Key             =   "编辑"
               Description     =   "编辑"
               Object.ToolTipText     =   "编辑病历标记图"
               Object.Tag             =   "编辑"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Description     =   "审核"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "审核"
               Description     =   "审核"
               Object.ToolTipText     =   "审核当前报告"
               Object.Tag             =   "审核"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "驳回"
               Key             =   "驳回"
               Description     =   "驳回"
               Object.ToolTipText     =   "驳回当前报告"
               Object.Tag             =   "驳回"
               ImageKey        =   "Rollback"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_41"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找病人病历"
               Object.Tag             =   "查找"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "显示"
               Key             =   "显示"
               Object.ToolTipText     =   "显示报告模板"
               Object.Tag             =   "显示"
               ImageKey        =   "History"
               Style           =   1
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "模板"
               Key             =   "模板"
               Object.ToolTipText     =   "将当前文本内容保存为报告模板"
               Object.Tag             =   "模板"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2715
      Left            =   4320
      TabIndex        =   21
      Top             =   3360
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
      TabIndex        =   23
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
   Begin MSComctlLib.ImageList iLstTab 
      Left            =   8040
      Top             =   5880
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
            Picture         =   "frmBillEdit.frx":B14A
            Key             =   "申请"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":B6E4
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":BC7E
            Key             =   "Template"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":C218
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":C7B2
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   195
      Left            =   1440
      TabIndex        =   24
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
      TabIndex        =   18
      Top             =   8760
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillEdit.frx":CD4C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5794
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picDoc 
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   8355
      TabIndex        =   25
      Top             =   1080
      Width           =   8415
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   420
         ScaleHeight     =   6495
         ScaleWidth      =   7515
         TabIndex        =   38
         Top             =   2040
         Width           =   7515
         Begin MSComctlLib.TreeView tvwElement 
            Height          =   1395
            Left            =   5850
            TabIndex        =   41
            Top             =   3135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   2461
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "iLstItem"
            Appearance      =   1
         End
         Begin zl9CISCore.ctrlPatientFile ProFile1 
            Height          =   5175
            Index           =   1
            Left            =   2160
            TabIndex        =   16
            Top             =   360
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   9128
            AllowEdit       =   -1  'True
            Border_Width    =   0
         End
         Begin zl9CISCore.ctrlPatientFile ProFile1 
            Height          =   5175
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   120
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   9128
            AllowEdit       =   -1  'True
            Border_Width    =   0
         End
      End
      Begin VB.PictureBox picAdvice 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   1365
         ScaleHeight     =   1815
         ScaleWidth      =   9255
         TabIndex        =   26
         Top             =   195
         Width           =   9255
         Begin VB.TextBox txt采集 
            Height          =   300
            Left            =   5040
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmd采集 
            Height          =   285
            Left            =   6600
            Picture         =   "frmBillEdit.frx":D5E0
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "选择检验标本"
            Top             =   350
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txt附加 
            Height          =   300
            Left            =   6440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk开始时间 
            BackColor       =   &H80000005&
            Caption         =   "要求时间"
            Height          =   225
            Left            =   315
            TabIndex        =   4
            ToolTipText     =   "是否安排时间"
            Top             =   420
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txt单量 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   12
            Top             =   1080
            Width           =   1380
         End
         Begin VB.TextBox txt频率 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1350
            TabIndex        =   10
            Top             =   1080
            Width           =   2500
         End
         Begin VB.TextBox txt总量 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4725
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1080
            Width           =   1380
         End
         Begin VB.CheckBox chk紧急 
            BackColor       =   &H80000005&
            Caption         =   "紧急(&J)"
            Height          =   225
            Left            =   7200
            TabIndex        =   8
            Top             =   405
            Width           =   945
         End
         Begin VB.CommandButton cmdExt 
            Height          =   285
            Left            =   8040
            Picture         =   "frmBillEdit.frx":D6D6
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "选择检验标本"
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   285
            Left            =   5280
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cbo执行科室 
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "frmBillEdit.frx":D7CC
            Left            =   1350
            List            =   "frmBillEdit.frx":D7CE
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1440
            Width           =   2500
         End
         Begin VB.TextBox txt医嘱内容 
            Height          =   300
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   3945
         End
         Begin VB.ComboBox cbo医生 
            Height          =   300
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1425
            Width           =   1590
         End
         Begin VB.TextBox txt医生嘱托 
            Height          =   300
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   9
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmd频率 
            Enabled         =   0   'False
            Height          =   240
            Left            =   3575
            Picture         =   "frmBillEdit.frx":D7D0
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(F4)"
            Top             =   1110
            Width           =   270
         End
         Begin MSComCtl2.DTPicker txt开始时间 
            Height          =   300
            Left            =   1350
            TabIndex        =   5
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   70778883
            CurrentDate     =   38022
         End
         Begin VB.Label lbl采集 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "采集方式"
            Height          =   180
            Left            =   4275
            TabIndex        =   40
            Top             =   420
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Line lineTitleSplit 
            BorderColor     =   &H80000000&
            X1              =   400
            X2              =   1440
            Y1              =   320
            Y2              =   320
         End
         Begin VB.Label lbl附加 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "检验标本"
            Height          =   180
            Left            =   5640
            TabIndex        =   39
            Top             =   45
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl单量 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "每次"
            Height          =   180
            Left            =   6660
            TabIndex        =   37
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl单量单位 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   8460
            TabIndex        =   36
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl频率 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "频率"
            Height          =   180
            Left            =   960
            TabIndex        =   35
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl总量单位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   6150
            TabIndex        =   34
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl总量 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "共"
            Height          =   180
            Left            =   4335
            TabIndex        =   33
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lbl执行科室 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行科室"
            Height          =   180
            Left            =   600
            TabIndex        =   32
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label lbl医嘱内容 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请项目"
            Height          =   180
            Left            =   600
            TabIndex        =   31
            Top             =   45
            Width           =   720
         End
         Begin VB.Label lbl开始时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "要求时间"
            Height          =   180
            Left            =   600
            TabIndex        =   30
            Top             =   435
            Width           =   720
         End
         Begin VB.Label lbl开嘱医生 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请医生"
            Height          =   180
            Left            =   5175
            TabIndex        =   29
            Top             =   1485
            Width           =   720
         End
         Begin VB.Label lbl医生嘱托 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生嘱托"
            Height          =   180
            Left            =   585
            TabIndex        =   28
            Top             =   795
            Width           =   720
         End
         Begin VB.Line lineSplit 
            X1              =   0
            X2              =   1080
            Y1              =   1800
            Y2              =   1800
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEdit_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Auditing 
         Caption         =   "审核报告(&A)"
      End
      Begin VB.Menu mnuEdit_Rollback 
         Caption         =   "驳回报告(&B)"
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
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "清空内容(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Template 
         Caption         =   "保存模板(&M)"
      End
   End
   Begin VB.Menu mnuOrder_1 
      Caption         =   "病历(&A)"
      Visible         =   0   'False
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
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Insert 
         Caption         =   "插入元素(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "删除元素(&D)"
         Visible         =   0   'False
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
      Begin VB.Menu mnuTemplate 
         Caption         =   "报告模板(&M)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPatientInformation 
         Caption         =   "申请项目(&I)"
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
Attribute VB_Name = "frmBillEdit"
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
Private bln护士站 As Boolean
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1
Private DeptID As Long '开单科室
Private mblnShow申请 As Boolean
Private PatientDate As Date '病人就诊或入院时间
Private AdviceID As Long, SendNO As Long '医嘱ID、发送号
Private sCheckNo As String '发送单据号
Private iRecordType As Integer '记录性质
Private alngFileID(1) As Long '申请和报告ID
Private intType As Integer '诊疗类别:-1=其他、0=检查组合、1=手术、2=中药、4=检验
Private iTabIndex As Integer
Private mlng前提ID As Long, bln医技执行 As Boolean
Private mblnMoved As Boolean
Private mstrPrivs As String

'医嘱编辑
Private strAdviceText As String '医嘱内容
Private str类别 As String, lngClinicID As Long, strClinicName As String, str标本部位 As String
Private strSequence As String, lng频率次数 As Long, lng频率间隔 As Long, str间隔单位 As String '频率
Private int计价特性 As Integer, int执行性质 As Integer, lng病人科室ID As Long
Private mstr性别 As String
Private mstrLike As String
Private gint过敏登记有效天数 As Integer
Private rsRelativeAdvice As ADODB.Recordset '相关医嘱
Private strExtData As String '附加项目

Private ifInitItem As Boolean '是否在进入申请时直接显示申请项目

Private iCurrElementIndex As Integer '当前元素顺序号
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Sub ShowMe(frmParent As Object, ByVal lng医嘱ID As Long, ByVal lng单据ID As Long, ByVal lng病历ID As Long, ByVal str医嘱内容 As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional ByVal blnMoved As Boolean = False)
'strPrivs：权限串。每一位代表一种权限，0－无该权限、1－有该权限。
'   第1位：审核报告
'   第2位：驳回报告
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String, tmpDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    Dim rsDept As New ADODB.Recordset, strDept As String, strDeptName As String
    Dim strSQL As String
    
    On Error Resume Next
    '初始化
    If blnMoved Then ReadOnly = True
    mblnMoved = blnMoved
    
    strSQL = "Select a.病人ID,a.主页ID,a.挂号单,Decode(a.主页ID,Null,0,1),b.ID,b.名称,a.医生嘱托," + _
        "医嘱内容,开始执行时间,紧急标志,执行频次,总给予量,单次用量,开嘱医生,nvl(b.计算单位,' ') As 计算单位,b.类别,nvl(a.标本部位,' ') As 标本部位,A.执行科室ID " + _
        "From 病人医嘱记录 a,诊疗项目目录 b Where (a.ID=[1] Or a.相关ID=[1]) And a.诊疗项目ID=b.ID Order By nvl(a.相关ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, lng医嘱ID)
    If rsTmp.EOF Then Unload Me: Exit Sub
    lngClinicID = rsTmp(4): strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    '构造附加项目串
    rsTmp.MoveNext
    If Not rsTmp.EOF Then
        If rsTmp!类别 = "C" Then lngClinicID = rsTmp(4) '检验项目
    End If
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)
        If rsTmp!类别 = "C" Then tmpDiagName = tmpDiagName & "," & rsTmp(5)
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    If Len(tmpDiagName) > 0 Then '检验项目
        strDiagName = Mid(tmpDiagName, 2)
        
        '置采集方式
        rsTmp.MoveFirst
        Me.cmd采集.Tag = rsTmp(4)
        Me.txt采集 = rsTmp(5): Me.txt采集.Tag = Me.txt采集
        
        rsTmp.MoveNext
    Else
        rsTmp.MoveFirst
    End If
    
    intType = -1
    Me.txt医嘱内容 = strDiagName
    If rsTmp!类别 = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "组合项目"), 0) = 1 Then
        '检查组合项目
        intType = 0
        Call AdviceSet检查手术(1, strExtData)
        txt医嘱内容.Text = Get检查手术名称(1, strDiagName)
        Me.txt附加 = Get部位名称
    ElseIf rsTmp!类别 = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intType = 1
        Call AdviceSet检查手术(2, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, strDiagName)
        Me.txt附加 = Get麻醉名称
    ElseIf InStr(",7,8,", rsTmp!类别) > 0 Then
        '中药配方(单味草药当配方处理)
        intType = 2
    ElseIf rsTmp!类别 = "C" Then
        '检验项目选择检验标本
        intType = 4
        Me.txt附加 = rsTmp("标本部位"): str标本部位 = rsTmp("标本部位")
        strExtData = strExtData & ";" & str标本部位
    End If
    
    alngFileID(0) = lng病历ID: PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = lng单据ID: bSample = False: AdviceID = lng医嘱ID
    
    '显示医嘱内容
    If IsNull(rsTmp("开始执行时间")) Then
        Me.chk开始时间.Visible = True: Me.lbl开始时间.Visible = False: Me.chk开始时间.Value = 0
        Me.txt开始时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"): Me.txt开始时间.Enabled = False
    Else
        Me.txt开始时间 = rsTmp("开始执行时间"): Me.txt开始时间.Enabled = True
    End If
    Me.chk紧急.Value = rsTmp("紧急标志")
    If Not IsNull(rsTmp("医生嘱托")) Then Me.txt医生嘱托 = rsTmp("医生嘱托")
    Me.txt频率 = rsTmp("执行频次"): Me.txt频率.Enabled = True: Me.cmd频率.Enabled = True
    Me.lbl总量单位.Caption = Trim(rsTmp("计算单位"))
    If Not IsNull(rsTmp("总给予量")) Then Me.txt总量 = rsTmp("总给予量"): Me.txt总量.Enabled = True
    If Not IsNull(rsTmp("单次用量")) Then Me.txt单量 = rsTmp("单次用量"): Me.txt单量.Enabled = True: Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = Trim(rsTmp("计算单位"))
    
    Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False
    strSQL = "Select 编码,名称 From 部门表 Where ID=[1]"
    Set rsDept = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("执行科室ID"), 0))
    If Not rsDept.EOF Then
        Me.cbo执行科室.AddItem rsDept("编码") & "-" & rsDept("名称")
        Me.cbo执行科室.Text = rsDept("编码") & "-" & rsDept("名称"): Me.cbo执行科室.Enabled = True
    End If
    Me.cbo医生.Clear: Me.cbo医生.AddItem rsTmp("开嘱医生")
    Me.cbo医生.Text = rsTmp("开嘱医生"): Me.cbo医生.Enabled = True
    Me.picAdvice.Enabled = False
    
    Me.stbThis.Panels(3).Visible = False: Me.stbThis.Panels(4).Visible = False
    
    If alngFileID(0) = 0 Then
        strSQL = "Select Count(*)" + _
            " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=1"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            MsgBox "未定义申请项目，不能编辑", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From 病人病历内容 Where 病历记录ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From 病历文件组成 Where 病历文件ID=" + FileTypeID + " And 填写时机=1"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    MsgBox "未定义申请项目，不能编辑", vbInformation, gstrSysName
                    Unload Me
                    Exit Sub
                End If
            Else
                MsgBox "没有申请内容，不能编辑", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        End If
    End If
    '初始化结束
    
    '判断能否编辑申请
    If Not ReadOnly Then
        '此处肯定不查询后备表
        strSQL = "Select 报告ID From 病人医嘱发送 Where 医嘱ID=[1] And Not 报告ID Is Null"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, AdviceID)
        If Not rsTmp.EOF Then ReadOnly = True
    End If
    bAllowEdit = Not ReadOnly
    
    iCurrElementIndex = 0

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
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "说明", "说明", 1800
    End With

    '获取可选元素列表
'    GetElementList
'    mnuOrder_Add_FileList
    '获取病人信息
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If bSample Then
        Me.Caption = "全文示范"
        stbInfo.Visible = False
    Else
        Me.Caption = "病历文件(申请)"
        stbInfo.Visible = True
        If alngFileID(0) > 0 Then
            strSQL = "Select 病历名称 From 病人病历记录 Where ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        Else
            strSQL = "Select 名称 From 病历文件目录 Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(申请)"
        
        strSQL = "Select Nvl(门诊号,0),Nvl(住院号,0),姓名,Nvl(性别,' '),Nvl(年龄,' '),nvl(b.名称,' ') As 科室,nvl(c.名称,' ') As 病区,当前床号," + IIf(PatientType = 0, "就诊时间 ", "入院时间 ") + _
            "From 病人信息 a,部门表 b,部门表 c Where 病人ID=[1] And a.当前科室ID=b.ID(+) And a.当前病区ID=c.ID(+)"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "无病人信息"
        Else
            PatientDate = rsTmp(8)
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "门诊号：" & rsTmp(0), "住院号：" & rsTmp(1))
                .Item(1).Text = "姓名：" & rsTmp(2) & "，性别：" & rsTmp(3) & "，年龄：" & rsTmp(4)
                
                mstr性别 = rsTmp(3)
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
'    With Me.stbAdvInfo.Panels
'        .Item(1).Text = "项目：" + strDiagName
'        .Item(2).Text = "医嘱内容：" + str医嘱内容
'    End With
'    Me.stbDrAdviceInfo.Panels(1).Text = "医生嘱托：" + strDrAdvice
    
    ProFile1(0).AllowEdit = bAllowEdit
    '处理菜单及工具栏
    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.tbrMain.Buttons("保存").Visible = bAllowEdit
    Me.mnuEdit_Clear.Visible = bAllowEdit
    
    iTabIndex = -1
    TabFile.Tabs.Clear
    TabFile.Tabs.Add , "申请", "申请(&S)", "申请"
    TabFile.Tabs("申请").Selected = True
    '处理Tab
    Me.TabFile.Visible = False: Me.ProFile1(1).Visible = False

    Set ParentForm = frmParent
    
    SetItemFormat
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
End Sub

Public Sub ShowMe_Report(frmParent As Object, ByVal strNO As String, ByVal int记录性质 As Integer, ByVal lng单据ID As Long, ByVal lng病历ID As Long, ByVal str医嘱内容 As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional ByVal lng前提ID As Long = 0, Optional ByVal If医技执行 As Boolean = False, Optional ByVal blnShow申请 As Boolean = True, Optional ByVal lng医嘱ID As Long = 0, Optional blnMoved As Boolean = False, Optional strPrivs As String = "00")
    
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String, tmpDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    Dim rsDept As New ADODB.Recordset, strDept As String, strDeptName As String
    Dim rsCapture As New ADODB.Recordset '采集方式记录
    Dim strSQL As String
    
    On Error Resume Next
    '初始化
    If blnMoved Then ReadOnly = True
    mblnMoved = blnMoved
    mstrPrivs = strPrivs
    
    If ReadOnly Then
        tvwElement.Visible = False
        tbrMain.Buttons("显示").Visible = False
        tbrMain.Buttons("Split_5").Visible = False
        mnuTemplate.Visible = False
    End If
    
    mblnShow申请 = blnShow申请
    
    picAdvice.Visible = mnuPatientInformation.Checked
'    If blnShow申请 = False Then tvwElement.Visible = blnShow申请
    
    strSQL = "Select a.病人ID,a.主页ID,a.挂号单,Decode(a.主页ID,Null,0,1),b.ID,b.名称,a.医生嘱托,a.ID,a.申请ID," + _
        "医嘱内容,开始执行时间,紧急标志,执行频次,总给予量,单次用量,开嘱医生,b.类别,nvl(a.标本部位,' ') As 标本部位,c.发送号,执行科室ID,a.相关ID " + _
        "From 病人医嘱记录 a,诊疗项目目录 b,病人医嘱发送 c Where" & _
        " c.NO=[1] And c.记录性质=[2]" & _
        IIf(lng医嘱ID = 0, "", " And (A.ID=[3] Or A.相关ID=[3])") & " And a.诊疗项目ID=b.ID And a.ID=c.医嘱ID Order By nvl(a.相关ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, strNO, int记录性质, lng医嘱ID)
    If rsTmp.EOF Then Unload Me: Exit Sub
    lngClinicID = rsTmp(4): strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    sCheckNo = strNO: iRecordType = int记录性质
        
    '构造附加项目串
'    If Not rsTmp!类别 = "C" Then rsTmp.MoveNext
    Do While Not rsTmp.EOF
        If rsTmp!类别 = "C" Then
            tmpDiagName = tmpDiagName & "," & rsTmp(5)
            strExtData = strExtData & "," & rsTmp(4)
        End If
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst
    If Len(tmpDiagName) > 0 Then '检验项目
        strDiagName = Mid(tmpDiagName, 2)
        If Not rsTmp!类别 = "C" Then rsTmp.MoveNext
        
        '置采集方式
        strSQL = "Select b.ID,b.名称 From 病人医嘱记录 a ,诊疗项目目录 b " & _
            "Where a.诊疗项目ID=b.ID and a.id=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsCapture = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("相关ID"), 0))
        If Not rsCapture.EOF Then
            Me.cmd采集.Tag = rsCapture(0)
            Me.txt采集 = rsCapture(1): Me.txt采集.Tag = Me.txt采集
        End If
    End If
     
    intType = -1
    Me.txt医嘱内容 = strDiagName
    If rsTmp!类别 = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "组合项目"), 0) = 1 Then
        '检查组合项目
        intType = 0
        Call AdviceSet检查手术(1, strExtData)
        txt医嘱内容.Text = Get检查手术名称(1, strDiagName)
        Me.txt附加 = Get部位名称
    ElseIf rsTmp!类别 = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intType = 1
        Call AdviceSet检查手术(2, strExtData)
        txt医嘱内容.Text = Get检查手术名称(2, strDiagName)
        Me.txt附加 = Get麻醉名称
    ElseIf InStr(",7,8,", rsTmp!类别) > 0 Then
        '中药配方(单味草药当配方处理)
        intType = 2
    ElseIf rsTmp!类别 = "C" Then
        '检验项目选择检验标本
        intType = 4
        Me.txt附加 = rsTmp("标本部位"): str标本部位 = rsTmp("标本部位")
        strExtData = strExtData & ";" & str标本部位
    End If
   
    alngFileID(0) = IIf(IsNull(rsTmp(8)), 0, rsTmp(8))
    alngFileID(1) = lng病历ID: PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    
    PatientType = rsTmp(3): FileTypeID = lng单据ID: bSample = False: AdviceID = rsTmp(7): SendNO = rsTmp("发送号")
    mlng前提ID = lng前提ID: bln医技执行 = If医技执行
    
    '显示医嘱内容
    If IsNull(rsTmp("开始执行时间")) Then
        Me.chk开始时间.Visible = True: Me.lbl开始时间.Visible = False: Me.chk开始时间.Value = 0
        Me.txt开始时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"): Me.txt开始时间.Enabled = False
    Else
        Me.txt开始时间 = rsTmp("开始执行时间"): Me.txt开始时间.Enabled = True
    End If
    Me.chk紧急.Value = rsTmp("紧急标志")
    If Not IsNull(rsTmp("医生嘱托")) Then Me.txt医生嘱托 = rsTmp("医生嘱托")
    Me.txt频率 = rsTmp("执行频次"): Me.txt频率.Enabled = True: Me.cmd频率.Enabled = True
    Me.lbl总量单位.Caption = Trim(rsTmp("计算单位"))
    If Not IsNull(rsTmp("总给予量")) Then Me.txt总量 = rsTmp("总给予量"): Me.txt总量.Enabled = True
    If Not IsNull(rsTmp("单次用量")) Then Me.txt单量 = rsTmp("单次用量"): Me.txt单量.Enabled = True: Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = Trim(rsTmp("计算单位"))
    
    Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False
    strSQL = "Select 编码,名称 From 部门表 Where ID=[1]"
    Set rsDept = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("执行科室ID"), 0))
    If Not rsDept.EOF Then
        Me.cbo执行科室.AddItem rsDept("编码") & "-" & rsDept("名称")
        Me.cbo执行科室.Text = rsDept("编码") & "-" & rsDept("名称"): Me.cbo执行科室.Enabled = True
    End If
    Me.cbo医生.Clear: Me.cbo医生.AddItem rsTmp("开嘱医生")
    Me.cbo医生.Text = rsTmp("开嘱医生"): Me.cbo医生.Enabled = True
    Me.picAdvice.Enabled = False
    
    Me.stbThis.Panels(3).Text = "报告人：" + UserInfo.姓名: Me.stbThis.Panels(4).Text = "时间：" + Format(zlDatabase.Currentdate, "yy-MM-dd HH:mm:ss")
    
    If alngFileID(0) = 0 Then
        strSQL = "Select Count(*)" + _
            " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=1"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            alngFileID(0) = -1 '没有申请项目
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From 病人病历内容 Where 病历记录ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=1"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    alngFileID(0) = -1 '没有申请项目
                End If
            Else
                alngFileID(0) = -1 '没有申请项目
            End If
        End If
    End If
    
    If alngFileID(1) = 0 Then
        strSQL = "Select Count(*)" + _
            " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=2"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            MsgBox "未定义报告项目，不能编辑", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From 病人病历内容 Where 病历记录ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(1))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=2"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    MsgBox "未定义报告项目，不能编辑", vbInformation, gstrSysName
                    Unload Me
                    Exit Sub
                End If
            Else
                MsgBox "没有报告内容，不能编辑", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        End If
    End If
    '初始化结束
    
    '判断能否编辑申请
    If Not ReadOnly Then
        strSQL = "Select 报告ID From 病人医嘱发送 Where 医嘱ID=[1] And Not 报告ID Is Null"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, AdviceID)
        If Not rsTmp.EOF Then
            bAllowEdit = False
        Else
            bAllowEdit = True
        End If
    Else
        bAllowEdit = False
    End If
    
    iCurrElementIndex = 0

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
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "说明", "说明", 1800
    End With

    '获取可选元素列表
'    GetElementList
'    mnuOrder_Add_FileList
    '获取病人信息
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If bSample Then
        Me.Caption = "全文示范"
        stbInfo.Visible = False
    Else
        Me.Caption = "病历文件(报告)"
        stbInfo.Visible = True
        If alngFileID(1) > 0 Then
            strSQL = "Select 病历名称 From 病人病历记录 Where ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(1))
        Else
            strSQL = "Select 名称 From 病历文件目录 Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(报告)"
        
        strSQL = "Select Nvl(门诊号,0),Nvl(住院号,0),姓名,Nvl(性别,' '),Nvl(年龄,' '),nvl(b.名称,' ') As 科室,nvl(c.名称,' ') As 病区,当前床号," + IIf(PatientType = 0, "就诊时间 ", "入院时间 ") + _
            "From 病人信息 a,部门表 b,部门表 c Where 病人ID=[1] And a.当前科室ID=b.ID(+) And a.当前病区ID=c.ID(+)"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "无病人信息"
        Else
            PatientDate = rsTmp(8)
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "门诊号：" & rsTmp(0), "住院号：" & rsTmp(1))
                .Item(1).Text = "姓名：" & rsTmp(2) & "，性别：" & rsTmp(3) & "，年龄：" & rsTmp(4)
                
                mstr性别 = rsTmp(3)
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
'    With Me.stbAdvInfo.Panels
'        .Item(1).Text = "项目：" + strDiagName
'        .Item(2).Text = "医嘱内容：" + str医嘱内容
'    End With
'    Me.stbDrAdviceInfo.Panels(1).Text = "医生嘱托：" + strDrAdvice
    
    '判断能否编辑
    ProFile1(0).AllowEdit = False ' bAllowEdit
    ProFile1(1).AllowEdit = Not ReadOnly
    '处理菜单及工具栏
    Me.mnuFileSave.Visible = Not ReadOnly: Me.mnuFileSplit(1).Visible = Not ReadOnly
    Me.tbrMain.Buttons("保存").Visible = Not ReadOnly
    Me.mnuEdit_Clear.Visible = Not ReadOnly
    
    iTabIndex = -1
    TabFile.Tabs.Clear
    If alngFileID(0) > -1 And mblnShow申请 Then TabFile.Tabs.Add , "申请", "申请(&S)", "申请"
    TabFile.Tabs.Add , "报告", "报告(&B)", "报告"
    TabFile.Tabs("报告").Selected = True
    '处理Tab
    Me.TabFile.Visible = True

    Set ParentForm = frmParent
    
    SetItemFormat
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
End Sub

Public Sub ShowMe_Request(frmParent As Object, ByVal lng病人ID As Long, ByVal var主页或挂号 As Variant, ByVal lng单据ID As Long, _
    ByVal b护士站 As Boolean, Optional ByVal ModalWindow As Boolean = True, Optional ByVal lng前提ID As Long = 0)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    Dim strSQL As String
    
    On Error Resume Next
    '初始化
    mblnMoved = False
    
    alngFileID(0) = 0: PatientID = lng病人ID: CheckID = CStr(var主页或挂号)
    PatientType = IIf(TypeName(var主页或挂号) = "String", 0, 1): FileTypeID = lng单据ID: bSample = False: AdviceID = 0
    bln护士站 = b护士站: mlng前提ID = lng前提ID
    
    Me.stbThis.Panels(3).Visible = False: Me.stbThis.Panels(4).Visible = False
        
    strSQL = "Select Count(*)" + _
        " From 病历文件组成 Where 病历文件ID=[1] And 填写时机=1"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
    If rsTmp(0) = 0 Then
        MsgBox "未定义申请项目，不能编辑", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    '初始化结束
    
    '判断能否编辑申请
    bAllowEdit = True
    
    iCurrElementIndex = 0

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
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1800
        .Add , "说明", "说明", 1800
    End With

    '获取可选元素列表
'    GetElementList
'    mnuOrder_Add_FileList
    '获取病人信息
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    Me.Caption = "病历文件(申请)"
    stbInfo.Visible = True
        
    strSQL = "Select 名称 From 病历文件目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
    If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(申请)"
    
    If PatientType = 0 Then
        strSQL = "Select Nvl(a.门诊号,0),Nvl(a.住院号,0),a.姓名,Nvl(a.性别,' '),Nvl(a.年龄,' '),nvl(b.名称,' ') As 科室,nvl(c.名称,' ') As 病区,a.当前床号,a.就诊时间,d.病人科室ID " + _
        "From 病人信息 a,部门表 b,部门表 c,病人费用记录 d Where a.病人ID=[1] And a.当前科室ID=b.ID(+) And a.当前病区ID=c.ID(+) And " + _
        "d.记录性质=4 And d.记录状态 In (1,3) And d.序号=1 And d.门诊标志=1 And d.病人id=a.病人id And d.标识号=a.门诊号"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
    Else
        strSQL = "Select Nvl(a.门诊号,0),Nvl(a.住院号,0),a.姓名,Nvl(a.性别,' '),Nvl(a.年龄,' '),nvl(b.名称,' ') As 科室,nvl(c.名称,' ') As 病区,a.当前床号,a.入院时间,d.出院科室ID " + _
        "From 病人信息 a,部门表 b,部门表 c,病案主页 d Where a.病人ID=[1] And a.当前科室ID=b.ID(+) And a.当前病区ID=c.ID(+) And " + _
        "d.主页ID=[2] And d.病人ID=a.病人ID"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID, CheckID)
    End If
    DeptID = UserInfo.部门ID
    If rsTmp.EOF Then
        stbInfo.Panels(1).Text = "无病人信息"
    Else
        PatientDate = rsTmp(8)
        lng病人科室ID = rsTmp(9)
        DeptID = rsTmp(9)
        With stbInfo.Panels
            .Item(4).Text = IIf(PatientType = 0, "门诊号：" & rsTmp(0), "住院号：" & rsTmp(1))
            .Item(1).Text = "姓名：" & rsTmp(2) & "，性别：" & rsTmp(3) & "，年龄：" & rsTmp(4)
            
            mstr性别 = rsTmp(3)
            If PatientType = 0 Then
                .Item(2).Visible = False: .Item(3).Visible = False
            Else
                .Item(2).Text = "科室：" & rsTmp(5)
                .Item(3).Text = "病区：" & rsTmp(6) & "，床号：" & NVL(rsTmp(7))
            End If
        End With
        
        Me.Caption = rsTmp(2) + "-" + Me.Caption
    End If
    
    ProFile1(0).AllowEdit = bAllowEdit
    '处理菜单及工具栏
    iTabIndex = -1
    TabFile.Tabs.Clear
    TabFile.Tabs.Add , "申请", "申请(&S)", "申请"
    TabFile.Tabs("申请").Selected = True
    '处理Tab
    Me.TabFile.Visible = False: Me.ProFile1(1).Visible = False

    '初始输入项
    Me.txt开始时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    '初始医生列表
    Call Get开嘱医生(CLng(PatientID), bln护士站, "", 0, Me.cbo医生, PatientType + 1)
    
    Set ParentForm = frmParent
    
    initForm
    If intType = 4 Then strExtData = ";"
    
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
End Sub

Private Sub initForm()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
        "From 诊疗项目目录 A,诊疗单据应用 B,诊疗项目别名 C Where A.ID=B.诊疗项目ID And A.ID=C.诊疗项目ID " + _
        "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN([1],3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And B.病历文件ID=[2] And 应用场合=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, PatientType + 1, FileTypeID)

    If rsTmp.EOF Then Exit Sub

    intType = -1
    If rsTmp!类别ID = "D" And zlCommFun.NVL(GetItemField(rsTmp!诊疗项目ID, "组合项目"), 0) = 1 Then
        '检查组合项目
        intType = 0
    ElseIf rsTmp!类别ID = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intType = 1
    ElseIf InStr(",7,8,", rsTmp!类别ID) > 0 Then
        '中药配方(单味草药当配方处理)
        intType = 2
    ElseIf rsTmp!类别ID = "C" Then
        '检验项目选择检验标本
        intType = 4
    End If
    
    rsTmp.MoveFirst: If rsTmp.RecordCount = 1 Then ifInitItem = True '因为只有一个项目，所以无需选择，进入申请时直接显示申请项目

    SetItemFormat
End Sub

Private Sub SetItemFormat()   '根据申请项目决定显示方式
    Select Case intType
        Case 0
            Me.lbl医嘱内容.Caption = "检查项目": Me.lbl附加.Caption = "检查部位": Me.cmdExt.ToolTipText = "选择检查部位"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lbl医嘱内容.Caption = "手术项目": Me.lbl附加.Caption = "麻醉方式": Me.cmdExt.ToolTipText = "选择麻醉方式"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 4
            Me.lbl医嘱内容.Caption = "检验项目": Me.lbl附加.Caption = "检验标本": Me.cmdExt.ToolTipText = "选择检验标本"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
            Me.lbl采集.Visible = True: Me.txt采集.Visible = True: Me.cmd采集.Visible = True
        Case Else
            Me.lbl附加.Visible = False: Me.txt附加.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub EnableEditMenu(ByVal bAllowEdit As Boolean)
    Dim i As Integer
    Dim strSinglePriv As String
    
'    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.mnuEdit.Visible = bAllowEdit
'    Me.mnuOrder_1.Visible = bAllowEdit
    For i = 1 To Me.tbrMain.Buttons.Count
        If Me.tbrMain.Buttons(i).Description = "编辑" Then Me.tbrMain.Buttons(i).Visible = bAllowEdit
    Next
    
    '处理打印、审核、驳回的权限
    strSinglePriv = Left(mstrPrivs, 1)
    mnuEdit_Auditing.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("审核").Visible = (strSinglePriv = 1)
    
    strSinglePriv = Mid(mstrPrivs, 2, 1)
    mnuEdit_Rollback.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("驳回").Visible = (strSinglePriv = 1)
    
    If mstrPrivs Like "00*" Then
        mnuEdit_0.Visible = False
        tbrMain.Buttons("Split_4").Visible = False
    End If
    
    strSinglePriv = Mid(mstrPrivs, 3, 1)
    mnuPreview.Visible = (strSinglePriv = 1)
    mnuPrint.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("预览").Visible = (strSinglePriv = 1)
    tbrMain.Buttons("打印").Visible = (strSinglePriv = 1)
    
    
'    '处理审核和驳回权限
'    If mstrPrivs Like "00*" Then
'        mnuEdit_0.Visible = False
'        mnuEdit_Auditing.Visible = False
'        mnuEdit_Rollback.Visible = False
'
'        tbrMain.Buttons("Split_4").Visible = False
'        tbrMain.Buttons("审核").Visible = False
'        tbrMain.Buttons("驳回").Visible = False
'    Else
'        mnuEdit_Auditing.Visible = (Mid(mstrPrivs, 1, 1) = 1)
'        mnuEdit_Rollback.Visible = (Mid(mstrPrivs, 2, 1) = 1)
'
'        tbrMain.Buttons("审核").Visible = (Mid(mstrPrivs, 1, 1) = 1)
'        tbrMain.Buttons("驳回").Visible = (Mid(mstrPrivs, 2, 1) = 1)
'    End If
End Sub

Private Sub cmd采集_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID))
    Me.txt采集.SetFocus
    If Not rsTmp Is Nothing Then
        Me.cmd采集.Tag = rsTmp("ID")
        Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub FileList_Click(Index As Integer)
    If MsgBox("加载病历示范后，当前病历内容将被覆盖！是否继续？", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在加载："
    ProFile1(iTabIndex).LoadSample CLng(FileList(Index).Tag), Me.prbRefresh
    ProFile1(iTabIndex).SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.TabFile.Visible Then
        If ProFile1(1).Tag = "" Then Exit Sub
        
        ProFile1(1).Tag = ""
        If alngFileID(0) > -1 Then
            Me.MousePointer = vbHourglass
            BeginShowProgress "正在加载申请："
            ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, Me.prbRefresh, mlng前提ID, , , mblnMoved
            If alngFileID(0) = 0 Then Call ProFile1(0).SetDiagItem(lngClinicID, str标本部位)
        End If
        Me.MousePointer = vbHourglass
        BeginShowProgress "正在加载报告："
        ProFile1(1).ShowFile IIf(alngFileID(1) = 0, "", CStr(alngFileID(1))), PatientID, CheckID, PatientType, FileTypeID, bSample, 2, Me.prbRefresh, mlng前提ID, AdviceID, SendNO, mblnMoved
        If alngFileID(1) = 0 Then Call ProFile1(1).SetDiagItem(lngClinicID, str标本部位)
        ProFile1(1).SetActiveElement 1
    Else
        If ProFile1(0).Tag = "" Then Exit Sub
        
        ProFile1(0).Tag = ""
        Me.MousePointer = vbHourglass
        BeginShowProgress "正在加载申请："
        ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, Me.prbRefresh, mlng前提ID, , , mblnMoved
        If alngFileID(0) = 0 Then Call ProFile1(0).SetDiagItem(lngClinicID, str标本部位)
        ProFile1(0).SetActiveElement 1
    End If
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
    
    If picAdvice.Enabled Then
        Me.txt医嘱内容.SetFocus
        If ifInitItem Then Call txt医嘱内容_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    '有关医嘱的参数
    mstrLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    '显示报告模板
    mnuTemplate.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告模板", "0"))
    mnuPatientInformation.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "申请信息", "0"))
    tbrMain.Buttons("显示").Value = IIf(mnuTemplate.Checked, tbrPressed, tbrUnpressed)
    Me.tvwElement.Visible = mnuTemplate.Checked
    
    '皮试结果有效时间
    gint过敏登记有效天数 = Val(GetSysParVal(2))
    
    '第一次触发Activate事件时要加载单据
    ProFile1(0).Tag = "Loading": ProFile1(1).Tag = "Loading"
    ProFile1(0).ifShowDiagItem = False: ProFile1(1).ifShowDiagItem = False
    
    '---------权限控制-------------
    'strPrivs = gstrPrivs
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim lngTxtWidth As Single
    Dim lngDistance As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.cbrMain.Visible, Me.cbrMain.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    lngDistance = 300
    
    On Error Resume Next
    With stbInfo
        .Left = 0: .Top = Me.cbrMain.Top + lngTools
        .Width = Me.ScaleWidth
        
        If PatientType = 0 Then
            .Panels(1).MINWIDTH = .Width - .Panels(4).MINWIDTH
        Else
            .Panels(1).MINWIDTH = 2 * (.Width - .Panels(4).MINWIDTH) / 5
            .Panels(2).MINWIDTH = 1.5 * (.Width - .Panels(4).MINWIDTH) / 5
            .Panels(3).MINWIDTH = 1.5 * (.Width - .Panels(4).MINWIDTH) / 5
        End If
    End With
    With picDoc
        .Left = 0: .Top = stbInfo.Top + stbInfo.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - lngStatus - IIf(TabFile.Visible, TabFile.Height, 0) - .Top
    End With
    With picAdvice
        .Left = 0: .Top = 0
        .Width = picDoc.ScaleWidth
    End With
    With lineSplit
        .X2 = picAdvice.Width + .X1
    End With
    With Me.chk紧急
        .Left = picAdvice.Width - Me.lbl开始时间.Left - .Width
        If .Left < Me.txt采集.Left + Me.txt采集.Width + lngDistance Then .Left = Me.txt采集.Left + Me.txt采集.Width + lngDistance
    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl开始时间.Left - Me.cmdSel.Width - Me.txt医嘱内容.Left - lngDistance - _
        Me.lbl附加.Width - Me.cmdExt.Width - 60) / 2
    With Me.txt医嘱内容
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl附加.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt附加
        .Left = Me.lbl附加.Left + Me.lbl附加.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txt医生嘱托
        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl开始时间.Left - Me.txt频率.Left - Me.txt频率.Width - _
        (Me.lbl总量单位.Width + Me.lbl总量.Width + lngDistance + 2 * 30) - _
        (Me.lbl单量单位.Width + Me.lbl单量.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl总量.Left = Me.txt频率.Left + Me.txt频率.Width + lngDistance
    With Me.txt总量
        .Left = Me.lbl总量.Left + Me.lbl总量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl总量单位.Left = Me.txt总量.Left + Me.txt总量.Width + 30
    Me.lbl单量.Left = Me.lbl总量单位.Left + Me.lbl总量单位.Width + lngDistance
    With Me.txt单量
        .Left = Me.lbl单量.Left + Me.lbl单量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl单量单位.Left = Me.txt单量.Left + Me.txt单量.Width + 30
    
    With Me.cbo医生
        .Left = Me.txt单量.Left
        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    Me.lbl开嘱医生.Left = Me.cbo医生.Left - Me.lbl开嘱医生.Width
    
    With picFile
        .Left = 0
        .Top = IIf(picAdvice.Visible, picAdvice.Top + picAdvice.Height, 0)
        .Width = picDoc.ScaleWidth
        .Height = picDoc.ScaleHeight - .Top
    End With
    With TabFile
        .Left = 0: .Top = Me.ScaleHeight - lngStatus - .Height
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ProFile1(iTabIndex).Modified And ProFile1(iTabIndex).AllowEdit Then
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        
        If Not Me.TabFile.Visible Then  '报告时，同时保存申请和报告
            If MsgBox("是否保存下达的申请", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                mnuFileSave_Click
            End If
        Else
            If Val(GetSetting("ZLSOFT", "公共模块\zl9Pacswork", "忽略结果阴阳性", 0)) = 1 Then
                If MsgBox("是否保存填写的报告", vbQuestion + vbYesNo, gstrSysName) = vbYes Then SaveFile
            Else
                SaveFile
            End If
        End If
    End If
'    zlCommFun.OpenIme False
    
    Call SaveWinState(Me, App.ProductName)
    '保存显示模板选项
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告模板", IIf(mnuTemplate.Checked, 1, 0))
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "申请信息", IIf(mnuPatientInformation.Checked, 1, 0))
    On Error Resume Next
    ParentForm.EditFile_UnLoad Me.hWnd  '告诉上级窗口编辑已关闭
    ProFile1(0).Release
    ProFile1(1).Release
End Sub

Private Sub lvwDemo_DblClick()
    If Me.lvwDemo.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwDemo
        ProFile1(iTabIndex).LoadElementSample iCurrElementIndex, Mid(.SelectedItem.Key, 2)
        
        .Visible = False
    End With
    
    ProFile1(iTabIndex).SetActiveElement iCurrElementIndex
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
        Me.MousePointer = vbHourglass
        BeginShowProgress "正在刷新："
        ProFile1(iTabIndex).InsertElement Mid(.SelectedItem.Key, 2), iCurrElementIndex, Me.prbRefresh
        Me.prbRefresh.Visible = False
        Me.MousePointer = vbDefault

        Me.stbThis.Panels(2).Text = ""
        
        .Visible = False
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

Private Sub mnuEdit_Auditing_Click()
'    If MsgBox("确认审核该报告吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo DBError
    If alngFileID(iTabIndex) = 0 Then
        If Not SaveFile Then Exit Sub
    Else
        If ProFile1(iTabIndex).Modified Then _
            If Not SaveFile Then Exit Sub
    End If
    
    
    Call ExeFinish(AdviceID, SendNO, False)
    Unload Me
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExeFinish(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnCancel As Boolean)
    Dim strSQL As String
    
    gcnOracle.BeginTrans
    On Error GoTo DBError
    If blnCancel Then
        strSQL = "ZL_病人医嘱执行_Cancel(" & lngAdviceID & "," & lngSendNO & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
        strSQL = "ZL_影像检查_STATE(" & lngAdviceID & "," & lngSendNO & ",5)"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        strSQL = "ZL_病人医嘱执行_Finish(" & lngAdviceID & "," & lngSendNO & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
        strSQL = "ZL_影像检查_STATE(" & lngAdviceID & "," & lngSendNO & ",6)"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "报告审核"
End Sub

Private Sub mnuEdit_Char_Click()
    frmSpecChar.Show vbModal, Me
    zlCommFun.OpenIme True
    If gblnOK Then SendKeys frmSpecChar.mstrChar
    Unload frmSpecChar
End Sub

Private Sub mnuEdit_Clear_Click()
    On Error Resume Next
    
    ProFile1(iTabIndex).ClearContent
    ProFile1(iTabIndex).SetActiveElement iCurrElementIndex
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
        strSQL = "Select a.ID,nvl(b.部件,' ') From 病人病历内容 a,病历元素目录 b Where a.元素编码=b.编码 And a.ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngContentID)
        
        If Not rsTmp.EOF Then ProFile1(iTabIndex).CopyElement iCurrElementIndex, rsTmp("ID"), rsTmp(1)
    End If
End Sub

Private Sub mnuEdit_Exchange_Click()
    If MsgBox("所见单内容将覆盖其文本段内容，是否继续", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    ProFile1(iTabIndex).ChangeToText iCurrElementIndex
    
    If Not Me.mnuEdit_Text.Checked Then
        mnuEdit_Text_Click
    Else
        If Not ProFile1(iTabIndex).ShowText(iCurrElementIndex, True) Then Me.mnuEdit_Text.Checked = False: Me.tbrMain.Buttons("文本").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuEdit_Map_Click()
    ProFile1(iTabIndex).EditElement iCurrElementIndex
End Sub

Private Sub mnuEdit_Rollback_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    If MsgBox("确认要驳回该报告吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
    On Error GoTo DBError
    strSQL = "Select Nvl(执行过程,0) As 执行过程 From 病人医嘱发送 Where 医嘱ID=[1] And 发送号=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, SendNO)
    If rsTmp.EOF Then Exit Sub
    
    If rsTmp(0) <> 6 Then
        strSQL = "ZL_影像检查_STATE(" & AdviceID & "," & SendNO & ",5)"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        Call ExeFinish(AdviceID, SendNO, True)
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Template_Click()
    If Len(Trim(ProFile1(iTabIndex).CurrentText(iCurrElementIndex))) = 0 Then
        MsgBox "该报告文本没有内容，不能存为模板。", vbInformation, gstrSysName
        Exit Sub
    End If
    frmBillSave.ShowMe Me, ProFile1(iTabIndex).ElementID(iCurrElementIndex), ProFile1(iTabIndex).CurrentText(iCurrElementIndex)
End Sub

Private Sub mnuEdit_Text_Click()
    If ProFile1(iTabIndex).ShowText(iCurrElementIndex, Not Me.mnuEdit_Text.Checked) Then Me.mnuEdit_Text.Checked = Not Me.mnuEdit_Text.Checked
    Me.tbrMain.Buttons("文本").Value = IIf(Me.mnuEdit_Text.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    Call SaveFile
End Sub
Private Function SaveFile() As Boolean
    Dim sTmpFileID As String
    Dim iMsgReturn As Integer
    
    SaveFile = False
    If Me.TabFile.Visible Then  '报告时，同时保存申请和报告
        If Val(GetSetting("ZLSOFT", "公共模块\zl9Pacswork", "忽略结果阴阳性", 0)) = 0 Then
            iMsgReturn = MsgBox("请确认报告结果是否为阳性？" & vbCrLf & "选择取消则放弃保存。", vbYesNoCancel + vbQuestion + vbDefaultButton1, gstrSysName)
            If iMsgReturn = vbCancel Then Exit Function
            iMsgReturn = IIf(iMsgReturn = vbYes, 1, 0)
        Else
            iMsgReturn = 0
        End If
        
        If alngFileID(0) > -1 And alngFileID(1) = 0 Then '要保存申请
            
            If mblnShow申请 Then
'                If MsgBox("报告保存时将同时保存申请，之后申请将不能修改！是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            End If
            
            '保存申请
            sTmpFileID = ProFile1(0).SaveFile
            If Len(sTmpFileID) > 0 Then
                alngFileID(0) = CLng(sTmpFileID)
                
                CommitData 0
            Else
                Exit Function
            End If
        End If
        '保存报告
        sTmpFileID = ProFile1(1).SaveFile
        If Len(sTmpFileID) > 0 Then
            alngFileID(1) = CLng(sTmpFileID)
            
            CommitData 1, iMsgReturn
            
            ProFile1(0).AllowEdit = False '不允许再编辑申请
            SaveFile = True: Exit Function
        Else
            Exit Function
        End If
    Else
        '保存申请
        
        If Me.picAdvice.Enabled Then
            If MsgBox("申请保存后系统将自动产生临时医嘱，" + Chr(13) + "申请项目将不能修改！是否要保存？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            If Not ValidAdvice Then Exit Function
            If Not SaveAdvice Then Exit Function
            
            Me.picAdvice.Enabled = False
        End If
        
        sTmpFileID = ProFile1(0).SaveFile
        If Len(sTmpFileID) > 0 Then
            alngFileID(0) = CLng(sTmpFileID)
            
            CommitData 0
            SaveFile = True: Exit Function
        Else
            Exit Function
        End If
    End If
End Function
'填写医嘱相关数据
Private Sub CommitData(ByVal iCommitType As Integer, Optional ByVal iCheckResult As Integer = -1)
    On Error GoTo DBError
    If iCommitType = 1 Then '报告
        If iCheckResult = -1 Then
            gcnOracle.Execute "ZL_诊疗单据_报告('" & sCheckNo & "'," & iRecordType & "," & alngFileID(1) & "," & _
                IIf(bln医技执行, 1, 0) & "," & AdviceID & ")", , adCmdStoredProc
        Else
            gcnOracle.Execute "ZL_诊疗单据_报告('" & sCheckNo & "'," & iRecordType & "," & alngFileID(1) & "," & _
                IIf(bln医技执行, 1, 0) & "," & AdviceID & "," & iCheckResult & ")", , adCmdStoredProc
        End If
    Else '申请
        gcnOracle.Execute "ZL_诊疗单据_申请(" & AdviceID & "," & alngFileID(0) & ")", , adCmdStoredProc
    End If
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub
'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If Len(Trim(strAdviceText)) = 0 Then
        ValidAdvice = False
        MsgBox "必须输入申请项目！", vbInformation, gstrSysName
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Len(Trim(strSequence)) = 0 Then
        ValidAdvice = False
        MsgBox "必须指定频率！", vbInformation, gstrSysName
        Me.txt频率.SetFocus: Exit Function
    End If
    If Not Check开始时间(CStr(Me.txt开始时间)) Then
        ValidAdvice = False
        Me.txt开始时间.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt总量)) = 0 Then
        ValidAdvice = False
        MsgBox "请输入总量！", vbInformation, gstrSysName
        Me.txt总量.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt单量)) = 0 And Me.txt单量.Enabled Then
        ValidAdvice = False
        MsgBox "请输入单量！", vbInformation, gstrSysName
        Me.txt单量.SetFocus: Exit Function
    End If
    If Val(Me.txt单量) > Val(Me.txt总量) Then
        ValidAdvice = False
        MsgBox "单量不能大于总量！", vbInformation, gstrSysName
        Me.txt总量.SetFocus: Exit Function
    End If
End Function
'保存医嘱
Private Function SaveAdvice() As Boolean
    On Error GoTo DBError
    SaveAdvice = True
    
    SaveAdviceData
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveAdvice = False
    SaveErrLog
End Function

Private Sub SaveAdviceData()
    Dim strSQL As String
    Dim lngAdviceID As Long, lngTmpID As Long
    Dim iMaxSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng开嘱科室ID As Long, strDoctor As String, i As Integer
    Dim str执行科室ID As String, str执行科室ID1 As String
    Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer, tmpint执行性质 As Integer
    Dim rsDept As ADODB.Recordset

    gcnOracle.BeginTrans
    On Error GoTo DBError
    
    lngAdviceID = zlDatabase.GetNextId("病人医嘱记录")
    strSQL = "Select Max(序号) From 病人医嘱记录 Where 病人ID=[1]" & _
        " And " & IIf(PatientType = 1, "主页ID=[2]", "挂号单=[2]")
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, CheckID)
    If IsNull(rsTmp(0)) Then
        iMaxSeq = 0
    Else
        iMaxSeq = rsTmp(0)
    End If
    
    lng开嘱科室ID = Get开嘱科室ID(Me.cbo医生.ItemData(Me.cbo医生.ListIndex), lng病人科室ID, PatientType + 1)
    i = InStr(Me.cbo医生.Text, "-")
    If i > 0 Then strDoctor = Mid(Me.cbo医生, i + 1)
    If Len(Me.cbo执行科室.Text) = 0 Then
        str执行科室ID = "NULL"
    Else
        str执行科室ID = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex)
    End If
    
    tmpstr类别 = str类别: tmplngClinicID = lngClinicID: tmpint计价特性 = int计价特性
    tmpint执行性质 = int执行性质
    If intType = 4 Then
        '检验项目将采集方式作为主医嘱
        strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
        If rsTmp.State = adStateOpen Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Me.cmd采集.Tag)
        tmpstr类别 = rsTmp("类别"): tmplngClinicID = rsTmp("ID"): tmpint计价特性 = NVL(rsTmp("计价性质"), 0)
        tmpint执行性质 = NVL(rsTmp("执行科室"), 0)
        '取采集方式的执行部门
        Set rsDept = GetExeDepart(rsTmp("ID"), PatientType + 1, DeptID)
        If rsDept Is Nothing Then
            str执行科室ID1 = "NULL"
        Else
            str执行科室ID1 = rsDept("ID")
        End If
    End If
    
    If intType <> 4 Then
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr类别 & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
            IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
            "'" & str标本部位 & "','" & strSequence & "'," & _
            IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
            IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
            "'" & str间隔单位 & "',NULL," & _
            tmpint计价特性 & "," & _
            str执行科室ID & "," & _
            tmpint执行性质 & "," & Me.chk紧急.Value & "," & _
            IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
            "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
            IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If
    '保存相关医嘱
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("病人医嘱记录")
            iMaxSeq = iMaxSeq + 1
            With rsRelativeAdvice
                strSQL = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
                    "0,1," & _
                    "1,'" & .Fields("类别") & "'," & _
                    .Fields("ID") & ",NULL,NULL," & _
                    IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
                    IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
                    "'" & Replace(.Fields("名称"), "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
                    "'" & IIf(intType = 4, str标本部位, .Fields("标本部位")) & "','" & strSequence & "'," & _
                    IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
                    IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
                    "'" & str间隔单位 & "',NULL," & _
                    .Fields("计价性质") & "," & _
                    str执行科室ID & "," & _
                    .Fields("执行科室") & "," & Me.chk紧急.Value & "," & _
                    IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
                    "NULL," & _
                    lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                    "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
                    IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
                gcnOracle.Execute strSQL, , adCmdStoredProc
                
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    
    If intType = 4 Then
        '检验申请的采集方式放到最后
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr类别 & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
            IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
            "'" & str标本部位 & "','" & strSequence & "'," & _
            IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
            IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
            "'" & str间隔单位 & "',NULL," & _
            tmpint计价特性 & "," & _
            str执行科室ID1 & "," & _
            tmpint执行性质 & "," & Me.chk紧急.Value & "," & _
            IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
            "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
            IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If

    gcnOracle.CommitTrans
    AdviceID = lngAdviceID
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "病人医嘱保存"
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuOrder_Delete_Click()
    Me.MousePointer = vbHourglass
    BeginShowProgress "正在刷新："
    ProFile1(iTabIndex).DeleteElement iCurrElementIndex, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Demo_Click()
    tbrMain_ButtonClick tbrMain.Buttons("元素")
End Sub

Private Sub mnuOrder_Insert_Click()
    tbrMain_ButtonClick tbrMain.Buttons("插入")
End Sub

Private Sub mnuPatientInformation_Click()
    Me.mnuPatientInformation.Checked = Not Me.mnuPatientInformation.Checked
    Me.picAdvice.Visible = Me.mnuPatientInformation.Checked
    Form_Resize
End Sub

Private Sub mnuPreview_Click()
    Dim frmPreview As frmCasePrint
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If alngFileID(iTabIndex) = 0 Then
        If MsgBox("该病历是新增的，打印之前系统将保存该份病历。是否继续", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1(iTabIndex).Modified Then _
            If MsgBox("打印之前是否保存该份病历", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
    End If
    If iTabIndex = 0 Then
        If bSample Then
            Set frmPreview = New frmCasePrint
            PrintOutCase Me, frmPreview, 0, True, 1, 0, alngFileID(iTabIndex), False, 0, 1
            frmPreview.Preview Me, 0, True, 1, 0, alngFileID(iTabIndex), False, 0, 1
        Else
            Set frmPreview = New frmCasePrint
            PrintOutCase Me, frmPreview, 5, True, -1 * CLng(Val(alngFileID(iTabIndex))), CLng(PatientID), CheckID, False, 0, 1
            frmPreview.Preview Me, 5, True, -1 * CLng(Val(alngFileID(iTabIndex))), CLng(PatientID), CheckID, False, 0, 1
        End If
    Else
        '打印报告
        PrintDiagReport AdviceID, SendNO, Me, 1, Me.picBuffer, mblnMoved
    End If
End Sub

Private Sub mnuPrint_Click()
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If alngFileID(iTabIndex) = 0 Then
        If MsgBox("该病历是新增的，打印之前系统将保存该份病历。是否继续", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1(iTabIndex).Modified Then _
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
    If iTabIndex = 0 Then
        If bSample Then
            PrintOutCase Me, Printer, 0, True, 1, 0, alngFileID(iTabIndex), False, 0, 1
        Else
            PrintOutCase Me, Printer, 5, True, -1 * CLng(Val(alngFileID(iTabIndex))), CLng(PatientID), CheckID, False, 0, 1
        End If
    Else
        '打印报告
        PrintDiagReport AdviceID, SendNO, Me, 2, Me.picBuffer, mblnMoved
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
    BeginShowProgress "正在加载："
    If iTabIndex = 0 Then
        ProFile1(iTabIndex).ShowFile IIf(alngFileID(iTabIndex) = 0, "", CStr(alngFileID(iTabIndex))), PatientID, CheckID, PatientType, FileTypeID, bSample, iTabIndex + 1, Me.prbRefresh, mlng前提ID, , , mblnMoved
    Else
        ProFile1(iTabIndex).ShowFile IIf(alngFileID(iTabIndex) = 0, "", CStr(alngFileID(iTabIndex))), PatientID, CheckID, PatientType, FileTypeID, bSample, iTabIndex + 1, Me.prbRefresh, mlng前提ID, AdviceID, SendNO, mblnMoved
    End If
    ProFile1(iTabIndex).SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuStatus_Click()
    Me.mnuStatus.Checked = Not Me.mnuStatus.Checked
    Me.stbThis.Visible = Me.mnuStatus.Checked
    Form_Resize
End Sub

Private Sub mnuTemplate_Click()
    mnuTemplate.Checked = Not mnuTemplate.Checked
    tbrMain.Buttons("显示").Value = IIf(mnuTemplate.Checked, tbrPressed, tbrUnpressed)
    tvwElement.Visible = mnuTemplate.Checked
    
    Call picFile_Resize
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
    With tvwElement
        .Left = 0: .Top = 0
        .Width = 3000: .Height = picFile.ScaleHeight
        .Width = IIf(tvwElement.Visible, 3000, 0)
    End With
    
    With ProFile1(iTabIndex)
        .Left = IIf(tvwElement.Visible, tvwElement.Left + tvwElement.Width, 0): .Top = 0
        .Width = picFile.ScaleWidth - .Left
        .Height = picFile.ScaleHeight
         
        If tvwElement.Visible Then
            If .Width + tvwElement.Width > picFile.ScaleWidth Then Me.Width = .Width + tvwElement.Width
            If .Height > picFile.ScaleHeight Then Me.Height = .Height + picFile.Top
        End If
    End With
End Sub

Private Sub ProFile1_ElementGotFocus(Index As Integer, ByVal ElementIndex As Integer, ByVal ElementType As Integer)
    If iCurrElementIndex <> ElementIndex And ProFile1(Index).AllowEdit Then
        ShowTemplate ProFile1(Index).ElementID(ElementIndex)
    End If
    
    iCurrElementIndex = ElementIndex
    If ProFile1(Index).AllowEdit Then
        EnableEditMenu True
        ShowEditMenu ElementType
    End If
End Sub

Private Sub ProFile1_Resize(Index As Integer)
    If Me.Width < ProFile1(Index).Width Then Me.Width = ProFile1(Index).Width
End Sub

Private Sub TabFile_Click()
    Select Case TabFile.SelectedItem.Key
        Case "申请"
            If iTabIndex = 0 Then Exit Sub
            
            iTabIndex = 0
        Case "报告"
            If iTabIndex = 1 Then Exit Sub
            
            iTabIndex = 1
    End Select
            
    Me.ProFile1(0).Visible = False
    Me.ProFile1(1).Visible = False
    picFile_Resize
    '设置编辑菜单
    If alngFileID(iTabIndex) > -1 Then
        EnableEditMenu ProFile1(iTabIndex).AllowEdit
    Else
        EnableEditMenu False
    End If
    Me.ProFile1(iTabIndex).Visible = True
    
    If Not ProFile1(iTabIndex).AllowEdit Then tvwElement.Nodes.Clear
    iCurrElementIndex = 0: ProFile1(iTabIndex).SetActiveElement 1
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
            Me.PopupMenu Me.mnuOrder_Add
        Case "元素"
            With Me.lvwDemo
                GetElementDemoList ProFile1(iTabIndex).ElementID(iCurrElementIndex)
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "删除"
            mnuOrder_Delete_Click
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
        Case "审核"
            mnuEdit_Auditing_Click
        Case "驳回"
            mnuEdit_Rollback_Click
        Case "显示"
            mnuTemplate_Click
        Case "模板"
            mnuEdit_Template_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
            mnuExit_Click
    End Select
End Sub

Private Sub ShowEditMenu(ElementType As Integer)
    If Not ProFile1(iTabIndex).AllowEdit Then Exit Sub
    Select Case ElementType
        Case 2 '所见单
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = True
            Me.tbrMain.Buttons("文本").Value = IIf(ProFile1(iTabIndex).IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("转储").Enabled = True
            Me.tbrMain.Buttons("符号").Enabled = True
            Me.tbrMain.Buttons("编辑").Enabled = False
            Me.tbrMain.Buttons("模板").Enabled = False
        Case 3 '标记图
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = False
            Me.tbrMain.Buttons("文本").Value = tbrUnpressed
            Me.tbrMain.Buttons("转储").Enabled = False
            Me.tbrMain.Buttons("符号").Enabled = False
            Me.tbrMain.Buttons("编辑").Enabled = True
            Me.tbrMain.Buttons("模板").Enabled = False
        Case 4 '专用纸
            Me.tbrMain.Buttons("复制").Enabled = False
            Me.tbrMain.Buttons("文本").Enabled = True
            Me.tbrMain.Buttons("文本").Value = IIf(ProFile1(iTabIndex).IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("转储").Enabled = True
            Me.tbrMain.Buttons("符号").Enabled = True
            Me.tbrMain.Buttons("编辑").Enabled = False
            Me.tbrMain.Buttons("模板").Enabled = False
        Case Else
            Me.tbrMain.Buttons("复制").Enabled = IIf(ElementType = 0, True, False)
            Me.tbrMain.Buttons("文本").Enabled = False
            Me.tbrMain.Buttons("文本").Value = tbrUnpressed
            Me.tbrMain.Buttons("转储").Enabled = False
            Me.tbrMain.Buttons("符号").Enabled = True
            Me.tbrMain.Buttons("编辑").Enabled = False
            Me.tbrMain.Buttons("模板").Enabled = IIf(ElementType = 0, True, False)
    End Select
    
    Me.mnuEdit_Copy.Enabled = Me.tbrMain.Buttons("复制").Enabled
    Me.mnuEdit_Char.Enabled = Me.tbrMain.Buttons("符号").Enabled
    Me.mnuEdit_Map.Enabled = Me.tbrMain.Buttons("编辑").Enabled
    Me.mnuEdit_Text.Enabled = Me.tbrMain.Buttons("文本").Enabled
    Me.mnuEdit_Text.Checked = IIf(Me.tbrMain.Buttons("文本").Value = tbrPressed, True, False)
    Me.mnuEdit_Exchange.Enabled = Me.tbrMain.Buttons("转储").Enabled
    Me.mnuEdit_Template.Enabled = Me.tbrMain.Buttons("模板").Enabled
    
    Me.mnuViewDoctor.Visible = Not bSample
End Sub

Private Sub GetElementList()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    Me.lvwItem.ListItems.Clear
    Err = 0: On Error GoTo ErrHand
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
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuOrder_Add_FileList()
    Dim rsFileList As New ADODB.Recordset
    Dim i As Integer, iNum As Integer
    Dim strSQL As String
    
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
            " 病历示范目录 Where ID=" & alngFileID(iTabIndex), Me.Caption
            
            FileTypeID = rsFileList(0)
        Else
            zlDatabase.OpenRecordset rsFileList, "Select 文件ID From" + _
            " 病人病历记录 Where ID=" & alngFileID(iTabIndex), Me.Caption
            
            FileTypeID = rsFileList(0)
        End If
    End If
    
    strSQL = "Select a.ID,a.名称 From 病历示范目录 a" + _
        " Where a.文件ID=[1] And a.类型=1" + _
        IIf(bSample, " And a.ID<>[2]", "") + _
        IIf(bSample, "", " And (a.科室ID=[3] Or" + _
        " a.科室ID Is Null)")
    Set rsFileList = OpenSQLRecord(strSQL, Me.Caption, FileTypeID, alngFileID(iTabIndex), UserInfo.部门ID)
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
    Dim strSQL As String
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select a.ID,a.名称,a.说明 From 病历示范目录 a" + _
        " Where a.元素ID=[1] And a.类型=2" + _
        IIf(bSample, "", " And (a.科室ID=[2] Or" + _
        " a.科室ID Is Null)")
    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, ElementID, UserInfo.部门ID)
    If rsTemp.EOF Then Exit Sub
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "元素": objItem.SmallIcon = "元素"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("说明").Index - 1) = IIf(IsNull(!说明), "", !说明)
            .MoveNext
        Loop
        Me.lvwDemo.ListItems(1).Selected = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub BeginShowProgress(ByVal strCaption As String)
    On Error Resume Next
    With prbRefresh
        .Left = stbThis.Panels(2).Left + Me.TextWidth(strCaption) + 200
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width + stbThis.Panels(2).Left - .Left
        
        stbThis.Panels(2).Text = strCaption
        .Visible = True: Me.Refresh
    End With
End Sub

'========以下是医嘱编辑==========

Private Sub cbo执行科室_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk紧急_Click()
    On Error Resume Next
    Me.txt医生嘱托.SetFocus
End Sub

Private Sub chk紧急_GotFocus()
    EnableEditMenu False
End Sub

Private Sub chk紧急_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk开始时间_Click()
    On Error Resume Next
    If Me.chk开始时间.Value = 1 Then
        Me.txt开始时间.Enabled = True: Me.txt开始时间.SetFocus
    Else
        Me.txt开始时间.Enabled = False
    End If
    
    If str类别 = "D" Then
        strAdviceText = Get检查手术内容(1, strClinicName)
    ElseIf str类别 = "F" Then
        strAdviceText = Get检查手术内容(2, strClinicName)
    End If
End Sub

Private Sub chk开始时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo医生_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ProFile1(iTabIndex).SetFocus
End Sub

Private Sub cmdExt_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim tmpExtData As String
    
    frmAdviceEditEx.mlngHwnd = Me.cbo医生.hWnd 'txt附加.Hwnd
    frmAdviceEditEx.mintType = IIf(intType = 4, 3, intType)
    frmAdviceEditEx.mint期效 = 1
    frmAdviceEditEx.mstr性别 = mstr性别
    If intType = 4 Then
        '检验项目
        frmAdviceEditEx.mlng项目ID = 0 'Split(strExtData, ";")(0)
        frmAdviceEditEx.mstrExtData = strExtData ' Split(strExtData, ";")(1)
    Else
        frmAdviceEditEx.mlng项目ID = lngClinicID
        frmAdviceEditEx.mstrExtData = strExtData
    End If
    frmAdviceEditEx.mint服务对象 = PatientType + 1

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me

    If Not frmAdviceEditEx.mblnOK Then
        zlControl.TxtSelAll Me.txt附加
        Me.txt附加.SetFocus
        Exit Sub
    Else
        tmpExtData = frmAdviceEditEx.mstrExtData
        If intType = 4 Then
            strExtData = Split(strExtData, ";")(0) + ";" + tmpExtData
        Else
            strExtData = tmpExtData
        End If
    End If
    Select Case intType
        Case 0 '检查组合部位
            Call AdviceSet检查手术(1, strExtData)
            strAdviceText = Get检查手术内容(1, strClinicName)
            Me.txt附加 = Get部位名称
        Case 1 '麻醉项目
            Call AdviceSet检查手术(2, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, strClinicName)
            strAdviceText = Get检查手术内容(2, strClinicName)
            Me.txt附加 = Get麻醉名称
        Case 4 '检验项目
            strAdviceText = strClinicName & "(" & tmpExtData & ")"
            Me.txt附加 = tmpExtData: str标本部位 = tmpExtData
    End Select
    txt附加.Tag = txt附加.Text
    Me.txt附加.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdExt_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset
    
    If intType = 4 Then
        '检验项目
        If LabsInput Then
            txt医嘱内容.Tag = txt医嘱内容.Text
            txt附加.Tag = txt附加.Text
            Me.txt医嘱内容.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            txt附加.Text = txt附加.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus
        End If
        Exit Sub
    End If
    
    With txt医嘱内容
        .Text = ""
        Set rsTmp = SelectDiagItem()
    End With
    
    If rsTmp Is Nothing Then '取消或无数据
        '恢复原值
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus: Exit Sub
    End If
    '新项目的录入
    
    '根据选择项目设置缺省医嘱信息
    If AdviceInput(rsTmp) Then
        '显示已缺省设置的值
        txt医嘱内容.Tag = txt医嘱内容.Text
        txt附加.Tag = txt附加.Text
        Me.txt医嘱内容.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '恢复原值
        txt医嘱内容.Text = txt医嘱内容.Tag
        txt附加.Text = txt附加.Tag
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus
    End If
End Sub

Private Sub cmdSel_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
        
    int范围 = 1
    strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
        " A.英文名称,A.频率次数,A.频率间隔,nvl(A.间隔单位,' ') As 间隔单位" & _
        " From 诊疗频率项目 A Where A.适用范围=" & int范围 & _
        " Order by A.编码"
    vRect = GetControlRect(txt频率.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "诊疗频率", , , , , , True, vRect.Left, vRect.Top, txt频率.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
        End If
        txt频率.Text = strSequence
        Call zlControl.TxtSelAll(txt频率)
        txt频率.SetFocus: Exit Sub
    End If
    Me.cmd频率.Tag = rsTmp("名称"): Me.txt频率 = Me.cmd频率.Tag: strSequence = Me.cmd频率.Tag
    lng频率次数 = rsTmp("频率次数"): lng频率间隔 = rsTmp("频率间隔"): str间隔单位 = Trim(rsTmp("间隔单位"))

    txt频率.SetFocus
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd频率_GotFocus()
    EnableEditMenu False
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Me.mnuToolbar, 2
End Sub

Private Sub tvwElement_DblClick()
    With tvwElement
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key Like "C*" Then Exit Sub
        
        ProFile1(iTabIndex).InsertTemplate iCurrElementIndex, .SelectedItem.Tag
    End With
End Sub

Private Sub tvwElement_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call tvwElement_DblClick
End Sub

Private Sub txt采集_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt采集)
End Sub

Private Sub txt采集_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt采集.Text = txt采集.Tag Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID), Me.txt采集)
    If Not rsTmp Is Nothing Then
        Me.cmd采集.Tag = rsTmp("ID")
        Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt采集_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt采集.Text <> txt采集.Tag Then
        txt采集.Text = txt采集.Tag
    End If
End Sub

Private Sub txt单量_GotFocus()
    EnableEditMenu False
    zlControl.TxtSelAll txt单量
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt单量_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt单量) Then Me.txt单量 = 1: Exit Sub
    Me.txt单量 = CInt(Me.txt单量)
    If CInt(Me.txt单量) < 1 Then Me.txt单量 = 1
End Sub

Private Sub txt附加_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt附加_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt附加)
End Sub

Private Sub txt附加_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt附加)
    End If
End Sub

Private Sub txt附加_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt附加.Text = txt附加.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        cmdExt_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt附加_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt附加.Text <> txt附加.Tag Then
        txt附加.Text = txt附加.Tag
    End If
End Sub

Private Sub txt开始时间_GotFocus()
    EnableEditMenu False
End Sub

Private Sub txt开始时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt开始时间_Validate(Cancel As Boolean)
    On Error Resume Next
    If Not Check开始时间(CStr(txt开始时间)) Then
        Cancel = True
        txt开始时间.SetFocus
    Else
        If str类别 = "D" Then
            strAdviceText = Get检查手术内容(1, strClinicName)
        ElseIf str类别 = "F" Then
            strAdviceText = Get检查手术内容(2, strClinicName)
        End If
    End If
End Sub

Private Sub txt频率_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt频率)
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmd频率.Tag <> "" And txt频率.Text = strSequence And txt频率.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt频率.Text = "" Then
            If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
        Else
            int范围 = 1 '可选频率
            strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
                " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位" & _
                " From 诊疗频率项目 A Where A.适用范围=" & int范围 & _
                " And (A.编码 Like '" & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.名称) Like '" & mstrLike & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.简码) Like '" & mstrLike & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.英文名称) Like '" & mstrLike & UCase(txt频率.Text) & "%')" & _
                " Order by A.编码"
            vRect = GetControlRect(txt频率.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "诊疗频率", , , , , , True, vRect.Left, vRect.Top, txt频率.Height, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                End If
                txt频率.Text = strSequence
                Call zlControl.TxtSelAll(txt频率)
                txt频率.SetFocus: Exit Sub
            End If
            Me.cmd频率.Tag = rsTmp("名称"): Me.txt频率 = Me.cmd频率.Tag: strSequence = Me.cmd频率.Tag
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt频率_Validate(Cancel As Boolean)
    If cmd频率.Tag <> "" And txt频率.Text <> strSequence Then
        txt频率.Text = strSequence
    End If
End Sub

Private Sub txt医生嘱托_GotFocus()
    EnableEditMenu False
End Sub

Private Sub txt医生嘱托_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt医生嘱托_Validate(Cancel As Boolean)
    On Error Resume Next
    If zlCommFun.ActualLen(txt医生嘱托.Text) > txt医生嘱托.MaxLength Then
        MsgBox "输入内容不过超过 " & txt医生嘱托.MaxLength \ 2 & " 个汉字或 " & txt医生嘱托.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt医生嘱托.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt医嘱内容_DblClick()
    If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
End Sub

Private Sub txt医嘱内容_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt医嘱内容)
    End If
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt医嘱内容.Text = "" Then cmdSel_Click: Exit Sub
        If txt医嘱内容.Text = txt医嘱内容.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        With txt医嘱内容
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
        '新项目的录入
        
        '根据选择项目设置缺省医嘱信息
        If AdviceInput(rsTmp) Then
            '显示已缺省设置的值
            txt医嘱内容.Tag = txt医嘱内容.Text
            txt附加.Tag = txt附加.Text
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            txt附加.Text = txt附加.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt医嘱内容.Text <> txt医嘱内容.Tag Then
        txt医嘱内容.Text = txt医嘱内容.Tag
    End If
End Sub

Private Sub txt总量_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(Me.txt总量)
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt总量_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt总量) Then Me.txt总量 = 1: Exit Sub
    Me.txt总量 = CInt(Me.txt总量)
    If CInt(Me.txt总量) < 1 Then Me.txt总量 = 1
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function Check开始时间(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法
'说明：
'1.开始时间不能小于病人的入院时间
'2.开始时间必须小于终止时间
'3.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
'4.补录的医嘱开始时间不能大于当前时间
    Dim strInDate As String
    
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
        
    strInDate = Format(PatientDate, "yyyy-MM-dd HH:mm")
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "医嘱的开始执行时间不能小于病人的" & IIf(PatientType = 0, "就诊", "入院") & "时间 " & strInDate & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
'    If IsDate(strEnd) Then
'        If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(strEnd, "yyyy-MM-dd HH:mm") Then
'            strMsg = "医嘱的开始执行时间必须小于执行终止时间。"
'            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    If DateDiff("n", CDate(strStart), zlDatabase.Currentdate) > 30 Then
        strMsg = "开始执行时间不能太早于当前时间。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check开始时间 = True
End Function

Private Function SelectDiagItem() As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
        "From 诊疗项目目录 A,诊疗单据应用 B,诊疗项目别名 C Where A.ID=B.诊疗项目ID And A.ID=C.诊疗项目ID " + _
        "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN(" & (PatientType + 1) & ",3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And (A.编码 Like '" + txt医嘱内容 + "%' Or Upper(A.名称) Like '" + mstrLike + txt医嘱内容 + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(txt医嘱内容) + "%') And B.病历文件ID=" & FileTypeID & " And 应用场合=" & (PatientType + 1)
            
    With txt医嘱内容
        Me.stbThis.Panels(2).Text = "请选择申请项目..."
        Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "选择申请项目", True, .Text, "", True, True, True, .Left + Me.picAdvice.Left + Me.Left, .Top + Me.picAdvice.Top + Me.Top, .Height, False, True)
        Me.stbThis.Panels(2).Text = ""
    End With
End Function

Private Function SelectCap(Optional ByVal lngItemID As Long = 0, Optional ByVal QryStr As String = "", Optional blnNotSelect As Boolean = False) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
    If Len(QryStr) > 0 Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A,诊疗项目别名 C,诊疗用法用量 D Where A.ID=C.诊疗项目ID And A.ID=D.用法ID" + _
            " And A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & (PatientType + 1) & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
            " And Nvl(A.执行频率,0) IN(0,1)" + _
            " And D.项目ID=" & lngItemID & _
            " And (A.编码 Like '" + QryStr + "%' Or Upper(A.名称) Like '" + mstrLike + QryStr + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(QryStr) + "%')"
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                "From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID" + _
                " And A.类别='E' And A.操作类型='6'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                " And A.服务对象 IN(" & (PatientType + 1) & ",3) And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)" + _
                " And (A.编码 Like '" + QryStr + "%' Or Upper(A.名称) Like '" + mstrLike + QryStr + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(QryStr) + "%')"
        End If
    Else
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
            " And A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & (PatientType + 1) & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
            " And Nvl(A.执行频率,0) IN(0,1)" + _
            " And D.项目ID=" & lngItemID
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                "From 诊疗项目目录 A Where " + _
                " A.类别='E' And A.操作类型='6'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                " And A.服务对象 IN(" & (PatientType + 1) & ",3) And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)"
        End If
    End If
    If blnNotSelect Then
        If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
        OpenRecord rsTmp, strSQL, Me.Caption
        If Not rsTmp.EOF Then Set SelectCap = rsTmp
    Else
        tmpRect = GetControlRect(Me.txt采集.hWnd)
        Set SelectCap = zlDatabase.ShowSelect(Me, strSQL, 0, "采集方式", True, , , , , True, _
            tmpRect.Left, tmpRect.Top, Me.txt采集.Height, , , True)
    End If
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean, i As Long
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String

    On Error GoTo errH

    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    txt医嘱内容.Text = rsInput!名称 '暂时显示

    '需要输入更多数据的一些项目
    '---------------------------------------------------------------------------------------------------------------
    intTmpType = -1
    If rsInput!类别ID = "D" And zlCommFun.NVL(GetItemField(rsInput!诊疗项目ID, "组合项目"), 0) = 1 Then
        '检查组合项目
        intTmpType = 0
        strHelpText = "检查部位"
    ElseIf rsInput!类别ID = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intTmpType = 1
        strHelpText = "附加手术及麻醉方式"
    ElseIf InStr(",7,8,", rsInput!类别ID) > 0 Then
        '中药配方(单味草药当配方处理)
        intTmpType = 2
    ElseIf rsInput!类别ID = "C" Then
        '检验项目选择检验标本
        intTmpType = 4
        strHelpText = "检验项目"
    End If

    If intTmpType <> -1 Then
        frmAdviceEditEx.mlngHwnd = Me.cbo执行科室.hWnd ' txt医嘱内容.Hwnd
        frmAdviceEditEx.mintType = intTmpType
        frmAdviceEditEx.mint期效 = 1
        frmAdviceEditEx.mstr性别 = mstr性别
        frmAdviceEditEx.mlng项目ID = IIf(intTmpType = 4, FileTypeID, rsInput!诊疗项目ID)
        frmAdviceEditEx.mstrExtData = IIf(intTmpType = 4, rsInput!诊疗项目ID & ";" & NVL(rsInput("标本部位")), "") '新输入项目
        frmAdviceEditEx.mint服务对象 = PatientType + 1

        On Error Resume Next
        Me.stbThis.Panels(2).Text = "请选择" + strHelpText + "..."
        frmAdviceEditEx.Show 1, Me
        Me.stbThis.Panels(2).Text = ""
        On Error GoTo errH

        If Not frmAdviceEditEx.mblnOK Then Exit Function
        If frmAdviceEditEx.mstrExtData = "" Or (Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" And rsInput!类别ID <> "F") Then Exit Function
        
        If rsInput!类别ID = "D" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '检查的组合部位行
            Call AdviceSet检查手术(1, strExtData)
            txt医嘱内容.Text = Get检查手术名称(1, rsInput!名称)
            strAdviceText = Get检查手术内容(1, rsInput!名称)
            Me.txt附加 = Get部位名称
        ElseIf rsInput!类别ID = "F" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '手术的附加手术及麻醉项目行
            Call AdviceSet检查手术(2, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, rsInput!名称)
            strAdviceText = Get检查手术内容(2, rsInput!名称)
            Me.txt附加 = Get麻醉名称
        ElseIf rsInput!类别ID = "C" And frmAdviceEditEx.mstrExtData <> "" Then
            '获取采集方式
            Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
            If rsTmp Is Nothing Then
                MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
                Exit Function
            End If
            Me.cmd采集.Tag = rsTmp("ID")
            Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
            
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '检验项目
            strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
                "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
                "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
                "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
                "From 诊疗项目目录 A,诊疗单据应用 B,诊疗项目别名 C Where A.ID=B.诊疗项目ID And A.ID=C.诊疗项目ID " + _
                "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                "And A.服务对象 IN([1],3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)" + _
                " And A.ID=[2] And B.病历文件ID=[3] And 应用场合=[1]"
            If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
            Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType + 1, Split(Split(strExtData, ";")(0), ",")(0), FileTypeID)
            
            Call AdviceSet检查手术(3, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, "")
            strAdviceText = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
            Me.txt附加 = Split(strExtData, ";")(1)
            str标本部位 = Me.txt附加
        End If
    Else
        str标本部位 = Trim(rsInput("标本部位"))
        txt医嘱内容.Text = txt医嘱内容.Text & "(" & str标本部位 & ")"
        strAdviceText = txt医嘱内容.Text
        
        '检查的组合部位行
        Call AdviceSet检查手术(1, "")
    End If
    
    '开始时间
    Me.txt开始时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("执行安排ID") = 1 Then
        Me.lbl开始时间.Visible = False: Me.chk开始时间.Visible = True: Me.chk开始时间.Value = 0
        Me.txt开始时间.Enabled = False
    Else
        Me.lbl开始时间.Visible = True: Me.chk开始时间.Visible = False
        Me.txt开始时间.Enabled = True
    End If
    
    '处理频率
    If rsInput("执行频率ID") = 1 Then
        Me.txt频率.Enabled = False: Me.txt频率 = "一次性": Me.cmd频率.Enabled = False
    Else
        Me.txt频率.Enabled = True: Me.txt频率 = "": Me.cmd频率.Enabled = True
    End If
    strSequence = Me.txt频率
    
    '总量
    Me.txt总量 = "1": Me.lbl总量单位.Caption = rsInput("计算单位")
    
    '单量
    If (rsInput("执行频率ID") = 0 And InStr(",1,2,", rsInput("计算方式ID")) > 0) _
                    Or InStr(",5,6,", rsInput("类别ID")) > 0 Then
        Me.txt单量.Enabled = True: Me.txt单量 = "": Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = rsInput("计算单位")
    Else
        Me.txt单量.Enabled = False: Me.txt单量 = "": Me.txt单量.BackColor = Me.BackColor: Me.lbl单量单位.Caption = "" ' rsInput("计算单位")
    End If
    
    '执行科室
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType + 1, DeptID)
    If rsTmp Is Nothing Then
        Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cbo执行科室.Clear
        Me.cbo执行科室.AddItem rsTmp("编码") & "-" & rsTmp("名称"): Me.cbo执行科室.ItemData(0) = rsTmp("ID"): Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    Else
        Me.cbo执行科室.Clear
        Do While Not rsTmp.EOF
            Me.cbo执行科室.AddItem rsTmp("编码") & "-" & rsTmp("名称"): Me.cbo执行科室.ItemData(Me.cbo执行科室.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = True: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    End If
    
    '开嘱医生
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '根据申请项目决定显示方式
    
    str类别 = rsInput("类别ID"): lngClinicID = rsInput("诊疗项目ID"): Call ProFile1(0).SetDiagItem(lngClinicID, str标本部位)
    int计价特性 = rsInput("计价性质ID"): int执行性质 = rsInput("执行科室ID"): strClinicName = IIf(intType = 4, Me.txt医嘱内容, rsInput("名称"))
    
    AdviceInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LabsInput() As Boolean
'功能：编辑检验项目
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean, i As Long
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String, rsInput As New ADODB.Recordset

    On Error GoTo errH
    
    intTmpType = 4
    strHelpText = "检验项目"

    frmAdviceEditEx.mlngHwnd = Me.cbo执行科室.hWnd ' txt医嘱内容.Hwnd
    frmAdviceEditEx.mintType = intTmpType
    frmAdviceEditEx.mint期效 = 1
    frmAdviceEditEx.mstr性别 = mstr性别
    frmAdviceEditEx.mlng项目ID = FileTypeID
    frmAdviceEditEx.mstrExtData = strExtData
    frmAdviceEditEx.mint服务对象 = PatientType + 1

    On Error Resume Next
    Me.stbThis.Panels(2).Text = "请选择" + strHelpText + "..."
    frmAdviceEditEx.Show 1, Me
    Me.stbThis.Panels(2).Text = ""
    On Error GoTo errH

    If Not frmAdviceEditEx.mblnOK Then Exit Function
    If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    Me.cmd采集.Tag = rsTmp("ID")
    Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
    
    strAdviceText = txt医嘱内容.Text
    strExtData = frmAdviceEditEx.mstrExtData

    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
        "From 诊疗项目目录 A,诊疗单据应用 B,诊疗项目别名 C Where A.ID=B.诊疗项目ID And A.ID=C.诊疗项目ID " + _
        "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN([1],3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And A.ID=[2] And B.病历文件ID=[3] And 应用场合=[1]"
    If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
    Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType + 1, Split(Split(strExtData, ";")(0), ",")(0), FileTypeID)
    
    Call AdviceSet检查手术(3, strExtData)
    txt医嘱内容.Text = Get检查手术名称(2, "")
    strAdviceText = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
    Me.txt附加 = Split(strExtData, ";")(1)
    str标本部位 = Me.txt附加
    
    '开始时间
    Me.txt开始时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("执行安排ID") = 1 Then
        Me.lbl开始时间.Visible = False: Me.chk开始时间.Visible = True: Me.chk开始时间.Value = 0
        Me.txt开始时间.Enabled = False
    Else
        Me.lbl开始时间.Visible = True: Me.chk开始时间.Visible = False
        Me.txt开始时间.Enabled = True
    End If
    
    '处理频率
    If rsInput("执行频率ID") = 1 Then
        Me.txt频率.Enabled = False: Me.txt频率 = "一次性": Me.cmd频率.Enabled = False
    Else
        Me.txt频率.Enabled = True: Me.txt频率 = "": Me.cmd频率.Enabled = True
    End If
    strSequence = Me.txt频率
    
    '总量
    Me.txt总量 = "1": Me.lbl总量单位.Caption = rsInput("计算单位")
    
    '单量
    If (rsInput("执行频率ID") = 0 And InStr(",1,2,", rsInput("计算方式ID")) > 0) _
                    Or InStr(",5,6,", rsInput("类别ID")) > 0 Then
        Me.txt单量.Enabled = True: Me.txt单量 = "": Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = rsInput("计算单位")
    Else
        Me.txt单量.Enabled = False: Me.txt单量 = "": Me.txt单量.BackColor = Me.BackColor: Me.lbl单量单位.Caption = "" ' rsInput("计算单位")
    End If
    
    '执行科室
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType + 1, DeptID)
    If rsTmp Is Nothing Then
        Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cbo执行科室.Clear
        Me.cbo执行科室.AddItem rsTmp("编码") & "-" & rsTmp("名称"): Me.cbo执行科室.ItemData(0) = rsTmp("ID"): Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    Else
        Me.cbo执行科室.Clear
        Do While Not rsTmp.EOF
            Me.cbo执行科室.AddItem rsTmp("编码") & "-" & rsTmp("名称"): Me.cbo执行科室.ItemData(Me.cbo执行科室.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = True: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    End If
    
    '开嘱医生
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '根据申请项目决定显示方式
    
    str类别 = rsInput("类别ID"): lngClinicID = rsInput("诊疗项目ID"): Call ProFile1(0).SetDiagItem(lngClinicID, str标本部位)
    int计价特性 = rsInput("计价性质ID"): int执行性质 = rsInput("执行科室ID"): strClinicName = IIf(intType = 4, Me.txt医嘱内容, rsInput("名称"))
    
    LabsInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '重新加入部位行或附加手术行及麻醉项目行
    If int类型 = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    ElseIf int类型 = 3 Then
        '处理检验项目
        strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get检查手术内容(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    strDate = IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "", Format(Me.txt开始时间, "yy年MM月dd日"))
    
    If rsRelativeAdvice Is Nothing Then
        If int类型 = 1 Then
            Get检查手术内容 = txtMainAdvice & IIf(Len(str标本部位) = 0, "", "(" & str标本部位 & ")"): Exit Function
        Else
            Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice & IIf(Len(str标本部位) = 0, "", "(" & str标本部位 & ")"): Exit Function
        End If
    End If
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If int类型 = 1 Then
            If Len(Trim(rsRelativeAdvice("标本部位"))) > 0 Then
                strTmp = strTmp & "," & rsRelativeAdvice("标本部位")
            End If
        ElseIf Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") = "G" Then
                str麻醉 = rsRelativeAdvice("名称")
            Else
                strTmp = strTmp & "," & rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If int类型 = 1 Then
        If strTmp <> "" Then
            Get检查手术内容 = txtMainAdvice & "(" & Mid(strTmp, 2) & ")"
        Else
            Get检查手术内容 = txtMainAdvice
        End If
    Else
        If strTmp <> "" Or str麻醉 <> "" Then
            If str麻醉 <> "" Then
                Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " ") & "在 " & str麻醉 & " 下行 " & txtMainAdvice
            Else
                Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice
            End If
            If strTmp <> "" Then
                Get检查手术内容 = Get检查手术内容 & " 及 " & Mid(strTmp, 2)
            End If
        Else
            Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice
        End If
    End If
End Function

Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function

Private Function Get麻醉名称() As String
    If rsRelativeAdvice Is Nothing Then Get麻醉名称 = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") = "G" Then
                Get麻醉名称 = rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get部位名称() As String
    If rsRelativeAdvice Is Nothing Then Get部位名称 = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("标本部位"))) > 0 Then
            Get部位名称 = Get部位名称 & "," & rsRelativeAdvice("标本部位")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get部位名称) > 0 Then Get部位名称 = Mid(Get部位名称, 2)
End Function

Private Function GetExeDepart(ByVal lngDiagItem As Long, ByVal iPatientType As Integer, Optional ByVal lngDepartID As Long = 0) As ADODB.Recordset
'功能：获取执行科室
'   iPatientType：病人类型 1=门诊、2=住院
'   lngDepartID：开单科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo DBError
    
    If lngDepartID = 0 Then lngDepartID = UserInfo.部门ID
    
    strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem)
    Select Case rsTmp("执行科室")
        Case 0, 1, 2 '0-无执行的叮嘱；1-病人所在科室；2-病人所在病区
            strSQL = "Select B.ID,B.编码,B.名称 From 病人信息 A,部门表 B Where " & _
                IIf(rsTmp("执行科室") = 1, "a.当前科室ID", "a.当前病区ID") & "=B.ID And A.病人ID=[1] Order by B.编码"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID)
        Case 3 '开单人所在科室
            strSQL = "Select B.ID,B.编码,B.名称 From 部门表 B Where B.ID=[1] Order by B.编码"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDepartID)
        Case 4 '指定科室
            strSQL = "Select Distinct B.ID,B.编码,B.名称 From 诊疗执行科室 A,部门表 B Where A.诊疗项目ID=[1]" & _
                " And A.开单科室ID=[2] And A.执行科室ID=B.ID Order by B.编码"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem, lngDepartID)
            '查询一般部门
            If rsTmp.EOF Then
                strSQL = "Select Distinct B.ID,B.编码,B.名称 From 诊疗执行科室 A,部门表 B Where A.诊疗项目ID=[1]" & _
                    " And 病人来源=[2] And A.执行科室ID=B.ID Order by B.编码"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem, iPatientType)
            End If
            If rsTmp.EOF Then
                strSQL = "Select Distinct B.ID,B.编码,B.名称 From 诊疗执行科室 A,部门表 B Where A.诊疗项目ID=[1]" & _
                    " And A.执行科室ID=B.ID Order by B.编码"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem)
            End If
        Case 5 '院外执行
            Exit Function
    End Select
    
    
    If Not rsTmp.EOF Then Set GetExeDepart = rsTmp
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupCount(lng组合ID As Long) As Long
'功能：获取组合项目中的项目数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(*) as NUM From 诊疗项目组合 Where 诊疗组合ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng组合ID)
    If Not rsTmp.EOF Then GetGroupCount = zlCommFun.NVL(rsTmp!NUM, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get缺省用法ID(int类型 As Integer) As Long
'功能：返回缺省的给药途径或中药煎法
'参数：int类型=2-给药途径,3-中药煎法,4-中药用法
'      str性别=病人性别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型=[1]" & _
        " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
        " Order by 编码"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int类型)
    If Not rsTmp.EOF Then Get缺省用法ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemField(ByVal lng项目ID As Long, ByVal strField As String) As Variant
'功能：获取指定诊疗项目的指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get开嘱医生(ByVal lng病人ID As Long, ByVal bln护士站 As Boolean, str缺省医生 As String, lng医生ID As Long, _
    Optional objCbo As Object, Optional ByVal int范围 As Integer = 2) As Boolean
'功能：获取可用的开嘱医生在指定的下拉框中
'参数：lng病人科室ID=病人所在科室ID
'      bln护士站=是否由护士代医生下医嘱
'      objCbo=要加入医生清单的下拉框
'      str缺省医生=缺省定位的医生,如果不传objCbo,则先优先定位,再返回缺省医生和医生ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
    
    If bln护士站 Then
        '病人所在科室的医生
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID=" & lng病人科室ID & _
            " Order by A.简码"
        '病人所在病区各科的医生
        strSQL = "Select Distinct 病区ID From 床位状况记录 Where 科室ID=" & lng病人科室ID
        strSQL = "Select Distinct 科室ID From 床位状况记录 Where 病区ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
        '全院住院科室的医生
        strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(" & int范围 & ",3)"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
    Else '医生下医嘱时,限制为只能为医生本人
        strSQL = "Select ID,编号,姓名,简码 From 人员表 Where ID=" & UserInfo.ID
    End If

    OpenRecord rsTmp, strSQL, "zlCISCore"
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln护士站 Then
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            ElseIf bln护士站 Then
                If str缺省医生 <> "" Then
                    '缺省医生(住院医师)优先
                    rsTmp.Filter = "姓名='" & str缺省医生 & "'"
                Else
                    '病人科室的医生优先
                    rsTmp.Filter = "部门ID=" & lng病人科室ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem zlCommFun.NVL(rsTmp!简码) & "-" & rsTmp!姓名
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!姓名 = str缺省医生 Then
                Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    End If
    Get开嘱医生 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get开嘱科室ID(ByVal lng医生ID As Long, ByVal lng病人科室ID As Long, Optional ByVal int范围 As Integer = 2) As Long
'功能：由医生确定开嘱科室
'参数：int范围=1-门诊,2-住院(缺省)
'说明：在医生所属科室范围内,优先顺序如下：
'      1、病人科室
'      2、服务于门诊/住院病人的科室且为默认科室
'      3、服务于门诊/住院病人的科室
'      4、默认科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 4) As Long
    
    '可能部门没有性质
    strSQL = "Select Distinct C.编码,A.部门ID,Nvl(A.缺省,0) as 缺省,Nvl(B.服务对象,0) as 服务对象" & _
        " From 部门人员 A,部门性质说明 B,部门表 C" & _
        " Where A.部门ID=C.ID And A.部门ID=B.部门ID(+) And A.人员ID=[1]" & _
        " Order by C.编码"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng医生ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!部门ID = lng病人科室ID Then
            arr科室ID(1) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 And rsTmp!缺省 = 1 Then
            arr科室ID(2) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 Then
            If arr科室ID(3) = 0 Then arr科室ID(3) = rsTmp!部门ID
        ElseIf rsTmp!缺省 = 1 Then
            arr科室ID(4) = rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr科室ID) To UBound(arr科室ID)
        If arr科室ID(i) <> 0 Then
            Get开嘱科室ID = arr科室ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowTemplate(ByVal lngElementID As Long)
'显示可用于当前元素的模板树
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim objCurrNode As MSComctlLib.Node
    
    On Error GoTo errH
    strSQL = "Select Distinct 0 As 末级,上级ID,ID,名称,'' As 内容,编码 From 病历模板分类" & _
        " Start With ID In" & _
        " (Select A.模板分类ID From 病历模板应用 A,病历模板分类 B Where A.模板分类ID=B.ID And 病历元素ID=[1] And " & _
        "(B.所属人员 Is Null Or B.所属人员='" & UserInfo.姓名 & "'))" & _
        " Connect By Prior 上级ID=ID" & _
        " Union All" & _
        " Select 1,a.分类ID,a.ID,a.名称,a.内容,a.编码 From 病历模板内容 a,病历模板应用 b,病历模板分类 c" & _
        " Where a.分类id=b.模板分类id And b.模板分类ID=c.ID And b.病历元素id=[1] And (c.所属人员 Is Null Or c.所属人员='" & UserInfo.姓名 & "') Order By 末级,编码"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngElementID)
    
    tvwElement.Nodes.Clear
    Do While Not rsTmp.EOF
        With tvwElement
            If IsNull(rsTmp("上级ID")) Then
                Set objCurrNode = .Nodes.Add(, , IIf(rsTmp("末级") = 0, "C", "T") & rsTmp("ID"), rsTmp("名称"), _
                    IIf(rsTmp("末级") = 0, "Close", "Template"), IIf(rsTmp("末级") = 0, "Open", "Template"))
                objCurrNode.Expanded = True
            Else
                Set objCurrNode = .Nodes.Add("C" & rsTmp("上级ID"), tvwChild, IIf(rsTmp("末级") = 0, "C", "T") & rsTmp("ID"), rsTmp("名称"), _
                    IIf(rsTmp("末级") = 0, "Close", "Template"), IIf(rsTmp("末级") = 0, "Open", "Template"))
            End If
            objCurrNode.Tag = NVL(rsTmp("内容"))
        End With
        
        rsTmp.MoveNext
    Loop
    If tvwElement.Nodes.Count > 0 Then tvwElement.Nodes(1).Expanded = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

