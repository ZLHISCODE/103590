VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMapEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "图形标注"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "frmMapEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imgCur 
      Left            =   7425
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":014A
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":02AC
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":05C6
            Key             =   "Earse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":08E0
            Key             =   "Text"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   542
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   8130
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2220
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   59
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "用于存放原始大小的标记图"
         Top             =   4680
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.PictureBox picOrig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   630
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "用作保存整幅图,以便取消时恢复"
         Top             =   4680
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picBuf 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "临时的缓冲作图区"
         Top             =   4680
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6195
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4155
         Width           =   255
      End
      Begin MSComCtl2.FlatScrollBar scrLR 
         Height          =   255
         Left            =   420
         TabIndex        =   8
         Top             =   4245
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         MousePointer    =   99
         MouseIcon       =   "frmMapEdit.frx":0A42
         Arrows          =   65536
         LargeChange     =   100
         Orientation     =   1245185
         SmallChange     =   3
      End
      Begin MSComCtl2.FlatScrollBar scrUD 
         Height          =   3915
         Left            =   6750
         TabIndex        =   7
         Top             =   375
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   6906
         _Version        =   393216
         MousePointer    =   99
         MouseIcon       =   "frmMapEdit.frx":0D5C
         LargeChange     =   100
         Orientation     =   1245184
         SmallChange     =   3
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3660
         Left            =   360
         ScaleHeight     =   242
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   379
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   255
         Width           =   5715
         Begin VB.TextBox txtTmp 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1455
            MultiLine       =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "用于求当前输入的行数"
            Top             =   2790
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.PictureBox picTxt 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   1425
            MousePointer    =   1  'Arrow
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "移动或双击设置字体"
            Top             =   240
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   1335
            MaxLength       =   250
            MouseIcon       =   "frmMapEdit.frx":1076
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   300
            Visible         =   0   'False
            Width           =   180
         End
      End
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   615
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4695
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   810
      Top             =   3975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar cbrStyle 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   5730
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   8130
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Caption1        =   " 样式 "
      Child1          =   "tbrStyle"
      MinHeight1      =   22
      Width1          =   165
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrStyle 
         Height          =   330
         Left            =   750
         TabIndex        =   3
         Top             =   30
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillColor"
               Object.ToolTipText     =   "填充颜色"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillNone"
               Description     =   "不填充"
               Object.ToolTipText     =   "不填充"
               ImageKey        =   "FillNone"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillAll"
               Description     =   "实心填充"
               Object.ToolTipText     =   "实心填充"
               ImageKey        =   "FillAll"
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillHsc"
               Description     =   "横线填充"
               Object.ToolTipText     =   "横线填充"
               ImageKey        =   "FillHsc"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillVsc"
               Description     =   "坚线填充"
               Object.ToolTipText     =   "坚线填充"
               ImageKey        =   "FillVsc"
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillHV"
               Description     =   "网络填充"
               Object.ToolTipText     =   "网格填充"
               ImageKey        =   "FillHV"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillR"
               Description     =   "右斜线填充"
               Object.ToolTipText     =   "右斜线填充"
               ImageKey        =   "FillR"
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillL"
               Description     =   "左斜线填充"
               Object.ToolTipText     =   "左斜线填充"
               ImageKey        =   "FillL"
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillLR"
               Description     =   "交叉线填充"
               Object.ToolTipText     =   "交叉线填充"
               ImageKey        =   "FillLR"
               Style           =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineColor"
               Object.ToolTipText     =   "线条颜色"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineAll"
               Description     =   "实线"
               Object.ToolTipText     =   "实线"
               ImageKey        =   "LineAll"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineDot"
               Description     =   "点虚线"
               Object.ToolTipText     =   "点虚线"
               ImageKey        =   "LineDot"
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineDash"
               Description     =   "长虚线"
               Object.ToolTipText     =   "长虚线"
               ImageKey        =   "LineDash"
               Style           =   2
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineDashDot"
               Description     =   "点划线"
               Object.ToolTipText     =   "点划线"
               ImageKey        =   "LineDashDot"
               Style           =   2
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineDashDot2"
               Description     =   "双点划线"
               Object.ToolTipText     =   "双点划线"
               ImageKey        =   "LineDashDot2"
               Style           =   2
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line1"
               Description     =   "线宽1点"
               Object.ToolTipText     =   "线宽1点"
               ImageKey        =   "Line1"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line2"
               Description     =   "线宽2点"
               Object.ToolTipText     =   "线宽2点"
               ImageKey        =   "Line2"
               Style           =   2
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line3"
               Description     =   "线宽3点"
               Object.ToolTipText     =   "线宽3点"
               ImageKey        =   "Line3"
               Style           =   2
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line4"
               Description     =   "线宽4点"
               Object.ToolTipText     =   "线宽4点"
               ImageKey        =   "Line4"
               Style           =   2
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line5"
               Description     =   "线宽5点"
               Object.ToolTipText     =   "线宽5点"
               ImageKey        =   "Line5"
               Style           =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   8130
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Caption1        =   "工具 "
      Child1          =   "tbrTool"
      MinWidth1       =   115
      MinHeight1      =   22
      Width1          =   115
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   330
         Left            =   645
         TabIndex        =   1
         Top             =   30
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Move"
               Description     =   "移动"
               Object.ToolTipText     =   "移动图片"
               ImageKey        =   "Move"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line"
               Description     =   "线条"
               Object.ToolTipText     =   "线条"
               ImageKey        =   "Line"
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MLine"
               Description     =   "折线"
               Object.ToolTipText     =   "折线"
               ImageKey        =   "MLine"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rect"
               Description     =   "矩形"
               Object.ToolTipText     =   "矩形"
               ImageKey        =   "Rect"
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MRect"
               Description     =   "多边形"
               Object.ToolTipText     =   "多边形"
               ImageKey        =   "MRect"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Circle"
               Description     =   "圆形"
               Object.ToolTipText     =   "圆形"
               ImageKey        =   "Circle"
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Earse"
               Description     =   "擦除"
               Object.ToolTipText     =   "擦除"
               ImageKey        =   "Earse"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Text"
               Description     =   "文字"
               Object.ToolTipText     =   "文字"
               ImageKey        =   "Text"
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stamp"
               Description     =   "点"
               Object.ToolTipText     =   "点"
               ImageKey        =   "Stamp"
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnDo"
               Description     =   "撤消"
               Object.ToolTipText     =   "撤消"
               ImageKey        =   "UnDo"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ReDo"
               Description     =   "恢复"
               Object.ToolTipText     =   "恢复"
               ImageKey        =   "ReDo"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ZoomIn"
               Description     =   "放大"
               Object.ToolTipText     =   "放大"
               ImageKey        =   "ZoomIn"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ZoomOut"
               Description     =   "缩小"
               Object.ToolTipText     =   "缩小"
               ImageKey        =   "ZoomOut"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ZoomNone"
               Description     =   "原始尺寸"
               Object.ToolTipText     =   "原始尺寸"
               ImageKey        =   "ZoomNone"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Clear"
               Description     =   "清除"
               Object.ToolTipText     =   "清除"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Description     =   "确认"
               Object.ToolTipText     =   "确认更改并退出"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "取消更改并退出"
               ImageKey        =   "Exit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   195
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":11C8
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1322
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":147C
            Key             =   "MLine"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":15D6
            Key             =   "Rect"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1730
            Key             =   "MRect"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":188A
            Key             =   "Circle"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":19E4
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1B3E
            Key             =   "UnDo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1C98
            Key             =   "ZoomIn"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1DF2
            Key             =   "ZoomOut"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":1F4C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":20A6
            Key             =   "FillNone"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2200
            Key             =   "FillAll"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":235A
            Key             =   "FillHsc"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":24B4
            Key             =   "FillVsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":260E
            Key             =   "FillHV"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2768
            Key             =   "FillR"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":28C2
            Key             =   "FillL"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2A1C
            Key             =   "FillLR"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2B76
            Key             =   "LineAll"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2CD0
            Key             =   "LineDot"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2E2A
            Key             =   "LineDash"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":2F84
            Key             =   "LineDashDot"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":30DE
            Key             =   "LineDashDot2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3238
            Key             =   "Line1"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3392
            Key             =   "Line2"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":34EC
            Key             =   "Line3"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3646
            Key             =   "Line4"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":37A0
            Key             =   "Line5"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":38FA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3A54
            Key             =   "Earse"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3BAE
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3D08
            Key             =   "ReDo"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3E62
            Key             =   "ZoomNone"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapEdit.frx":3FBC
            Key             =   "Stamp"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mobjMapItems As MapItems  '入/出:要编辑的标记图元素集
Public mstrMapName As String '入:当前病历项目名
Public mlngMapID As Long '入:标记图ID,用于保存窗体
Public mobjCaseMap As StdPicture '入:要标注的图片
Public mblnModi As Boolean '入:是否可以编辑

Private mintKey As Integer '顺序增加的不重复的关键字
Private mcolOper As Collection '用于Undo,ReDo操作的历史记录堆栈
Public mintOper As Integer '当前栈顶指针(0为空)

Private marrXY() As POINTAPI '折线或多边线的点集

Private mstrTool As String  '当前使用的工具
Private mintItem As Integer '当前选中的元素

Private msngScale As Single '当前操作比例
Private mlngOrgX As Long, mlngOrgY As Long '起始基点坐标
Private mlngTmpX As Long, mlngTmpY As Long '动态临时坐标
Private mblnDblClick As Boolean '是否双击
Private mintStampNo As Integer   '记录STAMP类型中数字的大小
Private lngColor(9) As Long

Private Sub Form_Activate()
    tbrTool.Top = 0: tbrTool.Left = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If txt.Visible Then
                Call FinishInput '完成并退出输入状态
            ElseIf tbrTool.Buttons("MLine").Value = tbrPressed And UBound(marrXY) >= 2 Then
                '取消画折线
                Call GetBufferAll '恢复原始图象
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            ElseIf tbrTool.Buttons("MRect").Value = tbrPressed And UBound(marrXY) >= 2 Then
                '取消画多边形
                Call GetBufferAll '恢复原始图象
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            End If
        Case vbKeyY
            If Shift = 2 And tbrTool.Buttons("ReDo").Visible Then Call tbrTool_ButtonClick(tbrTool.Buttons("ReDo"))
        Case vbKeyZ
            If Not txt.Visible And Shift = 2 And tbrTool.Buttons("UnDo").Visible Then Call tbrTool_ButtonClick(tbrTool.Buttons("UnDo"))
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call RestoreWinState(Me, App.ProductName, mlngMapID)
    
    picColor.BackColor = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "FillColor", &HFFFFFF)
    img16.ListImages.Add , , picColor.Image
    tbrStyle.Buttons("FillColor").Image = img16.ListImages.Count
    tbrStyle.Buttons("FillColor").Tag = picColor.BackColor
    
    picColor.BackColor = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "LineColor", vbRed)
    img16.ListImages.Add , , picColor.Image
    tbrStyle.Buttons("LineColor").Image = img16.ListImages.Count
    tbrStyle.Buttons("LineColor").Tag = picColor.BackColor
    
    '初始化
    Set picDraw.MouseIcon = imgCur.ListImages("Move").Picture
    picDraw.MousePointer = 99
    
    gblnOK = False
    mintOper = 0: Set mcolOper = New Collection
    msngScale = 1: glngPen = 0: glngBrush = 0
    mlngTmpX = 0: mlngTmpY = 0
    ReDim marrXY(0) '表示为空
    mstrTool = "Move": mintItem = 0
    mintKey = mobjMapItems.Count + 1
    
    Set mcolOper = New Collection
    
    '设置权限
    cbrStyle.Visible = mblnModi
    If Not mblnModi Then
        For i = 1 To tbrTool.Buttons.Count
            If InStr("Move;ZoomIn;ZoomOut;ZoomNone;Exit", tbrTool.Buttons(i).Key) = 0 Then
                tbrTool.Buttons(i).Visible = False
            End If
        Next
        Caption = "标记图" & IIf(mstrMapName <> "", " - " & mstrMapName, "")
    Else
        Caption = "图形标注" & IIf(mstrMapName <> "", " - " & mstrMapName, "")
    End If
    
    Call SetOperState
    
    '显示标记图内容
    Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
    
    lngColor(1) = RGB(0, 255, 130)
    lngColor(2) = RGB(255, 130, 0)
    lngColor(3) = RGB(255, 255, 0)
    lngColor(4) = RGB(255, 0, 130)
    lngColor(5) = RGB(130, 255, 0)
    lngColor(6) = RGB(40, 40, 255)
    lngColor(7) = RGB(0, 255, 255)
    lngColor(8) = RGB(130, 0, 255)
    lngColor(9) = RGB(255, 0, 0)
End Sub

Private Sub SetScrollBar()
'设置：设置滚动条的位置及滚动值
    Dim intUD As Integer, intLR As Integer
    Dim blnUD As Boolean, blnLR As Boolean
    
ReCalc:
    intUD = picDraw.Height - (picBack.ScaleHeight - IIf(scrLR.Visible, scrLR.Height, 0))
    intLR = picDraw.Width - (picBack.ScaleWidth - IIf(scrUD.Visible, scrUD.Width, 0))
    
    If intUD <= 0 And intLR <= 0 Then
        scrUD.Visible = False: scrLR.Visible = False: picTmp.Visible = False
    ElseIf intUD > 0 And intLR > 0 Then
        scrUD.Visible = True: scrLR.Visible = True: picTmp.Visible = True
    ElseIf intUD > 0 Then
        scrUD.Visible = True: scrLR.Visible = False: picTmp.Visible = False
        If Not blnUD Then blnUD = True: GoTo ReCalc
    ElseIf intLR > 0 Then
        scrLR.Visible = True: scrUD.Visible = False: picTmp.Visible = False
        If Not blnLR Then blnLR = True: GoTo ReCalc
    End If
    
    If scrLR.Visible Then
        scrLR.Max = intLR
        
        scrLR.Left = 0
        scrLR.Top = picBack.ScaleHeight - scrLR.Height
        scrLR.Width = picBack.ScaleWidth - IIf(scrUD.Visible, scrUD.Width, 0)
        scrLR.Refresh
        Call scrLR_Change
    Else
        picDraw.Left = (picBack.ScaleWidth - IIf(scrUD.Visible, scrUD.Width, 0) - picDraw.Width) / 2
    End If
    If scrUD.Visible Then
        scrUD.Max = intUD
        
        scrUD.Top = 0
        scrUD.Left = picBack.ScaleWidth - scrUD.Width
        scrUD.Height = picBack.ScaleHeight - (IIf(scrLR.Visible, scrLR.Height, 0))
        scrUD.Refresh
        Call scrUD_Change
    Else
        picDraw.Top = (picBack.ScaleHeight - IIf(scrLR.Visible, scrLR.Height, 0) - picDraw.Height) / 2
    End If
    If picTmp.Visible Then
        picTmp.Left = scrUD.Left
        picTmp.Top = scrLR.Top
    End If
    Me.Refresh
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picBack.Height = Me.ScaleHeight - cbrTool.Height - IIf(cbrStyle.Visible, cbrStyle.Height, 0)
    Call SetScrollBar
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gblnOK And mintOper > 0 Then
        If MsgBox("你确实要放弃对该标记图所作的所有改动吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    End If
    mblnModi = False
    
    Call SaveWinState(Me, App.ProductName, mlngMapID)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "FillColor", tbrStyle.Buttons("FillColor").Tag
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "LineColor", tbrStyle.Buttons("LineColor").Tag
End Sub

Private Sub picDraw_DblClick()
    mblnDblClick = True
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strXY As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    Dim intText As Integer
    
    mblnDblClick = False
    mlngOrgX = x: mlngOrgY = y
    If Button = 1 Then
        If tbrTool.Buttons("Line").Value = tbrPressed Then
            '线条
            Call SetDrawStyleFromFace(msngScale)
            MoveToEx picDraw.hDC, x, y, 0
        ElseIf tbrTool.Buttons("Rect").Value = tbrPressed Then
            '矩形
            Call SetDrawStyleFromFace(msngScale)
        ElseIf tbrTool.Buttons("Circle").Value = tbrPressed Then
            '(椭)圆
            Call SetDrawStyleFromFace(msngScale)
        ElseIf tbrTool.Buttons("MLine").Value = tbrPressed Then
            '折线
            If UBound(marrXY) = 0 Then '折线中开始画的第一根线
                Call SetBufferAll '备份整个图象,以便取消本次本图
                Call SetDrawStyleFromFace(msngScale)
                MoveToEx picDraw.hDC, x, y, 0
                ReDim marrXY(1 To 1): marrXY(1).x = x: marrXY(1).y = y
            ElseIf UBound(marrXY) >= 2 Then
                '相同点不处理
                If marrXY(UBound(marrXY)).x = x And marrXY(UBound(marrXY)).y = y Then Exit Sub
                '折线中间段线的确认
                ReDim Preserve marrXY(1 To UBound(marrXY) + 1)
                marrXY(UBound(marrXY)).x = x: marrXY(UBound(marrXY)).y = y
                MoveToEx picDraw.hDC, marrXY(UBound(marrXY) - 1).x, marrXY(UBound(marrXY) - 1).y, 0
                LineTo picDraw.hDC, x, y
                picDraw.Refresh '必须刷新
                mlngTmpX = 0: mlngTmpY = 0 '之后第一次作图不需要恢复
            End If
        ElseIf tbrTool.Buttons("MRect").Value = tbrPressed Then
            '多边形
            If UBound(marrXY) = 0 Then '多边形中开始画的第一根线
                Call SetBufferAll '备份整个图象,以便取消本次本图
                Call SetDrawStyleFromFace(msngScale)
                MoveToEx picDraw.hDC, x, y, 0
                ReDim marrXY(1 To 1): marrXY(1).x = x: marrXY(1).y = y
            ElseIf UBound(marrXY) >= 2 Then
                If Abs(marrXY(1).x - x) <= (gcurPenWidth + 3) * msngScale _
                    And Abs(marrXY(1).y - y) <= (gcurPenWidth + 3) * msngScale _
                    And UBound(marrXY) >= 3 Then '点击第一点则自动完成
                    
                    '从缓冲中恢复区域
                    If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                        Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
                    End If
                    
                    Polygon picDraw.hDC, marrXY(1), UBound(marrXY)
                    picDraw.Refresh '必须刷新
                    
                    '加入集合时用原始尺寸
                    strXY = GetStrFromArr
                    Call GetRect(strXY, X1, Y1, X2, Y2) '最小的范围
                    
                    mintKey = mintKey + 1
                    mobjMapItems.Add 4, "", "", strXY, X1, Y1, X2, Y2, gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
                    Call NewOper(mobjMapItems.Count, 1)
                    
                    Call ResetDrawStyle
                    ReDim marrXY(0)
                    mlngTmpX = 0: mlngTmpY = 0
                Else '多边形中间段线的确认
                    '相同点不处理
                    If marrXY(UBound(marrXY)).x = x And marrXY(UBound(marrXY)).y = y Then Exit Sub
                    ReDim Preserve marrXY(1 To UBound(marrXY) + 1)
                    marrXY(UBound(marrXY)).x = x: marrXY(UBound(marrXY)).y = y
                    MoveToEx picDraw.hDC, marrXY(UBound(marrXY) - 1).x, marrXY(UBound(marrXY) - 1).y, 0
                    LineTo picDraw.hDC, x, y
                    picDraw.Refresh '必须刷新
                    mlngTmpX = 0: mlngTmpY = 0 '之后第一次作图不需要恢复
                End If
            End If
        ElseIf tbrTool.Buttons("Earse").Value = tbrPressed And mintItem > 0 Then
            '擦除图象
            Call DrawItemState(mintItem, False)
            
            Call NewOper(mintItem, 2)
            mobjMapItems.Remove mintItem
            
            Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
            mintItem = 0
        ElseIf tbrTool.Buttons("Text").Value = tbrPressed Then
            '文字工具
            If txt.Visible Then Call FinishInput
            Me.Refresh
            intText = inText(x, y)
            If intText > 0 Then
                '字体,字号,字色,0000；四位分别表示粗体,斜体,下线,删除线
                With mobjMapItems(intText)
                    txt.FontName = Split(.字体, ",")(0)
                    txt.FontSize = Split(.字体, ",")(1) * msngScale
                    txt.ForeColor = Split(.字体, ",")(2)
                    txt.FontBold = Mid(Split(.字体, ",")(3), 1, 1) = "1"
                    txt.FontItalic = Mid(Split(.字体, ",")(3), 2, 1) = "1"
                    txt.FontUnderline = Mid(Split(.字体, ",")(3), 3, 1) = "1"
                    txt.FontStrikethru = Mid(Split(.字体, ",")(3), 4, 1) = "1"
                                        
                    txt.Text = .内容
                    
                    txt.Left = .X1 * msngScale
                    txt.Top = .Y1 * msngScale
                    txt.Width = (.X2 - .X1) * msngScale
                    txt.Height = (.Y2 - .Y1) * msngScale
                End With
                
                '编辑时先删除原文本显示,完成后再重新显示
                Call NewOper(intText, 2)
                mobjMapItems.Remove intText
                '这句引起慢
                Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
            Else
                If txt.Tag = "" Then
                    txt.FontName = "宋体"
                    txt.FontSize = 9 * msngScale
                    txt.ForeColor = 0
                    txt.FontBold = False
                    txt.FontItalic = False
                    txt.FontUnderline = False
                    txt.FontStrikethru = False
                Else
                    txt.FontName = Split(txt.Tag, ",")(0)
                    txt.FontSize = Split(txt.Tag, ",")(1) * msngScale
                    txt.ForeColor = Split(txt.Tag, ",")(2)
                    txt.FontBold = Mid(Split(txt.Tag, ",")(3), 1, 1) = "1"
                    txt.FontItalic = Mid(Split(txt.Tag, ",")(3), 2, 1) = "1"
                    txt.FontUnderline = Mid(Split(txt.Tag, ",")(3), 3, 1) = "1"
                    txt.FontStrikethru = Mid(Split(txt.Tag, ",")(3), 4, 1) = "1"
                End If
                
                txt.Text = ""
                
                txt.Top = y: txt.Left = x
                Call GetTxtSize(txt, txt.Text, X1, Y1)
                txt.Width = X1 + 10
                txt.Height = Y1 + 6
            End If
            picTxt.Top = txt.Top - picTxt.Height / 2
            picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
            txt.Visible = True
            picTxt.Visible = True
            txt.SetFocus
            
            Call SetOperState
        ElseIf tbrTool.Buttons("Stamp").Value = tbrPressed Then '带数字的圆点
            
            Dim iHalfWidth As Integer
            iHalfWidth = 7
            '设置圆的画笔，颜色
            mintStampNo = mintStampNo + 1
            
            tbrStyle.Buttons("FillAll").Value = tbrPressed
            tbrStyle.Buttons("LineAll").Value = tbrPressed
            tbrStyle.Buttons("LineColor").Tag = vbBlue
            tbrStyle.Buttons("Line1").Value = tbrPressed
            tbrStyle.Buttons("FillColor").Tag = lngColor((mintStampNo Mod 9) + 1)
                'RGB(20 + 20 * (mintStampNo) * ((mintStampNo + 1) Mod 5), _
                            '80 + 20 * (mintStampNo) * ((mintStampNo + 2) Mod 4), 150 + 20 * (mintStampNo) * ((mintStampNo) Mod 2))
            'RGB(lngR + mintStampNo * 10, lngG + mintStampNo * 10, lngB + mintStampNo * 10)
            Call SetDrawStyleFromFace(msngScale)
            
            '画椭圆和文字
            Ellipse picDraw.hDC, mlngOrgX - iHalfWidth, mlngOrgY - iHalfWidth, mlngOrgX + iHalfWidth, mlngOrgY + iHalfWidth
            TextOut picDraw, Trim(str(mintStampNo)), mlngOrgX - iHalfWidth, mlngOrgY - iHalfWidth, mlngOrgX + iHalfWidth, mlngOrgY + iHalfWidth, _
                    "宋体,9,0,1000", msngScale
            picDraw.Refresh '必须刷新
            
            Call ResetDrawStyle
            
            '加入集合时用原始尺寸,文本是0，圆是5
            mintKey = mintKey + 1
            mobjMapItems.Add 5, "", "", "Stamp", mlngOrgX - iHalfWidth / msngScale, mlngOrgY - iHalfWidth / msngScale, _
                    CLng(mlngOrgX + iHalfWidth / msngScale), CLng(mlngOrgY + iHalfWidth / msngScale), _
                    gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, _
                    gcurPenWidth, "_" & mintKey
            Call NewOper(mobjMapItems.Count, 1)
            
            '加入集合时用原始尺寸,文本是0，圆是5
            mintKey = mintKey + 1
            mobjMapItems.Add 0, Trim(str(mintStampNo)), "宋体,9,0,1000", "Stamp", mlngOrgX - iHalfWidth / msngScale, mlngOrgY - iHalfWidth / msngScale, _
                    CLng(mlngOrgX + iHalfWidth / msngScale), CLng(mlngOrgY + iHalfWidth / msngScale), _
                    gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, _
                    gcurPenWidth, "_" & mintKey
            Call NewOper(mobjMapItems.Count, 1)
        End If
        
    End If
End Sub

Private Sub GetTxtSize(objMain As Object, strText As String, Optional ByRef W As Long, Optional ByRef H As Long, Optional ByRef h2 As Long)
'功能：返回文本框当前内容的合适尺寸
'返回：w,h整个尺寸,h2单行高度
    With objMain
        picTmp.FontName = .FontName
        picTmp.FontSize = .FontSize
        picTmp.FontBold = .FontBold
        picTmp.FontItalic = .FontItalic
        picTmp.FontUnderline = .FontUnderline
        picTmp.FontStrikethru = .FontStrikethru
        If strText = "" Then
            W = picTmp.TextWidth("AA")
            H = picTmp.TextHeight("A")
        Else
            W = picTmp.TextWidth(strText & "A")
            H = picTmp.TextHeight(strText)
        End If
        h2 = picTmp.TextHeight("A")
    End With
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngOX As Long, lngOY As Long
    Dim i As Integer, j As Integer
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    Dim arrTmp() As String, blnSnap As Boolean
    
    If Button = 1 Then
        If tbrTool.Buttons("Move").Value = tbrPressed Then
            '移动图片
            If scrLR.Visible Then
                lngOX = x - mlngOrgX
                If scrLR.Value - lngOX > scrLR.Max Then
                    scrLR.Value = scrLR.Max
                ElseIf scrLR.Value - lngOX < scrLR.Min Then
                    scrLR.Value = scrLR.Min
                Else
                    scrLR.Value = scrLR.Value - lngOX
                End If
            End If
            If scrUD.Visible Then
                lngOY = y - mlngOrgY
                If scrUD.Value - lngOY > scrUD.Max Then
                    scrUD.Value = scrUD.Max
                ElseIf scrUD.Value - lngOY < scrUD.Min Then
                    scrUD.Value = scrUD.Min
                Else
                    scrUD.Value = scrUD.Value - lngOY
                End If
            End If
        ElseIf tbrTool.Buttons("Line").Value = tbrPressed Then
            '线条
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY, msngScale)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(mlngOrgX, mlngOrgY, x, y, msngScale)
            MoveToEx picDraw.hDC, mlngOrgX, mlngOrgY, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh '必须刷新
            
            mlngTmpX = x: mlngTmpY = y
        ElseIf tbrTool.Buttons("Rect").Value = tbrPressed Then
            '矩形
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY, msngScale)
            End If
            mlngTmpX = x: mlngTmpY = y
            If Shift = 2 Then '画正方形
                Call ForceSquare(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY, msngScale)
            Rectangle picDraw.hDC, mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY
            picDraw.Refresh '必须刷新
        ElseIf tbrTool.Buttons("Circle").Value = tbrPressed Then
            '(椭)圆
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY, msngScale)
            End If
            mlngTmpX = x: mlngTmpY = y
            If Shift = 2 Then '画圆
                Call ForceSquare(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY, msngScale)
            Ellipse picDraw.hDC, mlngOrgX, mlngOrgY, mlngTmpX, mlngTmpY
            picDraw.Refresh '必须刷新
        ElseIf tbrTool.Buttons("MLine").Value = tbrPressed Then
            '第一段线(必须按下)或中间线段(可以不按)
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, x, y, msngScale)
            MoveToEx picDraw.hDC, marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh '必须刷新
            
            mlngTmpX = x: mlngTmpY = y
        ElseIf tbrTool.Buttons("MRect").Value = tbrPressed Then
            '第一段线(必须按下)或中间线段(可以不按)
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, x, y, msngScale)
            MoveToEx picDraw.hDC, marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh '必须刷新
            
            mlngTmpX = x: mlngTmpY = y
        End If
    Else
        If tbrTool.Buttons("MLine").Value = tbrPressed And UBound(marrXY) >= 2 Then
            '折线(中间其它线段的虚拟显示)
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, x, y, msngScale)
            MoveToEx picDraw.hDC, marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh '必须刷新
            
            mlngTmpX = x: mlngTmpY = y
        ElseIf tbrTool.Buttons("MRect").Value = tbrPressed And UBound(marrXY) >= 2 Then
            '多边形(中间其它线段的虚拟显示)
            '从缓冲中恢复区域
            If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
            End If
            '先保存区域到缓冲中
            Call SetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, x, y, msngScale)
            MoveToEx picDraw.hDC, marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh '必须刷新
            
            mlngTmpX = x: mlngTmpY = y
        Else
            mlngTmpX = 0: mlngTmpY = 0
        End If
        
        '显示捕捉图象
        blnSnap = False
        If tbrTool.Buttons("Earse").Value = tbrPressed Then
            For i = 1 To mobjMapItems.Count
                With mobjMapItems(i)
                    Select Case .类型
                        Case 0 '文本
                            
                        Case 1 '线条
                            If InLine(x, y, .X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, .线宽) Then
                                If mintItem <> i Then
                                    If mintItem > 0 Then Call DrawItemState(mintItem, False)
                                    mintItem = i: Call DrawItemState(i, True)
                                End If
                                blnSnap = True: Exit For
                            End If
                        Case 2 '折线
                            arrTmp = Split(.点集, ";")
                            For j = 0 To UBound(arrTmp) - 1
                                X1 = Split(arrTmp(j), ",")(0)
                                Y1 = Split(arrTmp(j), ",")(1)
                                X2 = Split(arrTmp(j + 1), ",")(0)
                                Y2 = Split(arrTmp(j + 1), ",")(1)
                                If InLine(x, y, X1 * msngScale, Y1 * msngScale, X2 * msngScale, Y2 * msngScale, .线宽) Then
                                    If mintItem <> i Then
                                        If mintItem > 0 Then Call DrawItemState(mintItem, False)
                                        mintItem = i: Call DrawItemState(i, True)
                                    End If
                                    blnSnap = True: Exit For
                                End If
                            Next
                        Case 3 '矩形
                            If InLine(x, y, .X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y1 * msngScale, .线宽) _
                                Or InLine(x, y, .X1 * msngScale, .Y1 * msngScale, .X1 * msngScale, .Y2 * msngScale, .线宽) _
                                Or InLine(x, y, .X1 * msngScale, .Y2 * msngScale, .X2 * msngScale, .Y2 * msngScale, .线宽) _
                                Or InLine(x, y, .X2 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, .线宽) Then
                                If mintItem <> i Then
                                    If mintItem > 0 Then Call DrawItemState(mintItem, False)
                                    mintItem = i: Call DrawItemState(i, True)
                                End If
                                blnSnap = True: Exit For
                            End If
                        Case 4 '多边形
                            arrTmp = Split(.点集, ";")
                            For j = 0 To UBound(arrTmp)
                                X1 = Split(arrTmp(j), ",")(0)
                                Y1 = Split(arrTmp(j), ",")(1)
                                If j = UBound(arrTmp) Then
                                    X2 = Split(arrTmp(0), ",")(0)
                                    Y2 = Split(arrTmp(0), ",")(1)
                                Else
                                    X2 = Split(arrTmp(j + 1), ",")(0)
                                    Y2 = Split(arrTmp(j + 1), ",")(1)
                                End If
                                If InLine(x, y, X1 * msngScale, Y1 * msngScale, X2 * msngScale, Y2 * msngScale, .线宽) Then
                                    If mintItem <> i Then
                                        If mintItem > 0 Then Call DrawItemState(mintItem, False)
                                        mintItem = i: Call DrawItemState(i, True)
                                    End If
                                    blnSnap = True: Exit For
                                End If
                            Next
                        Case 5 '(椭)圆
                            If InEllipse(x, y, .X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, .线宽) Then
                                If mintItem <> i Then
                                    If mintItem > 0 Then Call DrawItemState(mintItem, False)
                                    mintItem = i: Call DrawItemState(i, True)
                                End If
                                blnSnap = True: Exit For
                            End If
                    End Select
                End With
            Next
            '清除捕捉图象
            If Not blnSnap And mintItem > 0 Then Call DrawItemState(mintItem, False): mintItem = 0
        End If
    End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngX As Long, lngY As Long, strXY As String
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    
    If Button = 1 Then
        If tbrTool.Buttons("Line").Value = tbrPressed Then
            '线条
            MoveToEx picDraw.hDC, mlngOrgX, mlngOrgY, 0
            LineTo picDraw.hDC, x, y
            picDraw.Refresh
            
            Call ResetDrawStyle
            
            '加入集合时用原始尺寸
            mintKey = mintKey + 1
            mobjMapItems.Add 1, "", "", "", mlngOrgX / msngScale, mlngOrgY / msngScale, CLng(x / msngScale), CLng(y / msngScale), gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
            Call NewOper(mobjMapItems.Count, 1)
        ElseIf tbrTool.Buttons("Rect").Value = tbrPressed Then
            '矩形
            lngX = x: lngY = y
            If Shift = 2 Then '画正方形
                Call ForceSquare(mlngOrgX, mlngOrgY, lngX, lngY)
            End If
            Rectangle picDraw.hDC, mlngOrgX, mlngOrgY, lngX, lngY
            picDraw.Refresh '必须刷新
            
            Call ResetDrawStyle
            '加入集合时用原始尺寸
            mintKey = mintKey + 1
            mobjMapItems.Add 3, "", "", "", mlngOrgX / msngScale, mlngOrgY / msngScale, CLng(lngX / msngScale), CLng(lngY / msngScale), gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
            Call NewOper(mobjMapItems.Count, 1)
        ElseIf tbrTool.Buttons("Circle").Value = tbrPressed Then
            '(椭)圆
            lngX = x: lngY = y
            If Shift = 2 Then '画圆
                Call ForceSquare(mlngOrgX, mlngOrgY, lngX, lngY)
            End If
            Ellipse picDraw.hDC, mlngOrgX, mlngOrgY, lngX, lngY
            picDraw.Refresh '必须刷新
            
            Call ResetDrawStyle
            '加入集合时用原始尺寸
            mintKey = mintKey + 1
            mobjMapItems.Add 5, "", "", "", mlngOrgX / msngScale, mlngOrgY / msngScale, CLng(lngX / msngScale), CLng(lngY / msngScale), gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
            Call NewOper(mobjMapItems.Count, 1)
        ElseIf tbrTool.Buttons("MLine").Value = tbrPressed Then
            If mblnDblClick And UBound(marrXY) >= 2 Then
                '完成折线
                If marrXY(UBound(marrXY)).x <> x Or marrXY(UBound(marrXY)).y <> y Then
                    '最后两点相同则当作一点
                    ReDim Preserve marrXY(1 To UBound(marrXY) + 1)
                    marrXY(UBound(marrXY)).x = x: marrXY(UBound(marrXY)).y = y
                End If
                MoveToEx picDraw.hDC, marrXY(UBound(marrXY) - 1).x, marrXY(UBound(marrXY) - 1).y, 0
                LineTo picDraw.hDC, x, y
                picDraw.Refresh '必须刷新
                
                '加入集合时用原始尺寸
                strXY = GetStrFromArr
                Call GetRect(strXY, X1, Y1, X2, Y2) '最小的范围
                mintKey = mintKey + 1
                mobjMapItems.Add 2, "", "", strXY, X1, Y1, X2, Y2, gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
                Call NewOper(mobjMapItems.Count, 1)
                
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            ElseIf UBound(marrXY) = 1 Then
                If marrXY(1).x = x And marrXY(1).y = y Then
                    '相同点取消作图
                    Call GetBufferAll '恢复原始图象
                    Call ResetDrawStyle
                    ReDim marrXY(0)
                    mlngTmpX = 0: mlngTmpY = 0
                    Exit Sub
                End If
                '折线(第一段线的确认)
                MoveToEx picDraw.hDC, mlngOrgX, mlngOrgY, 0
                LineTo picDraw.hDC, x, y
                picDraw.Refresh
                ReDim Preserve marrXY(1 To 2): marrXY(2).x = x: marrXY(2).y = y
                mlngTmpX = 0: mlngTmpY = 0 '之后第一次作图不需要恢复
            End If
        ElseIf tbrTool.Buttons("MRect").Value = tbrPressed Then
            If mblnDblClick And UBound(marrXY) >= 3 Then '至少有两条边三个点才能完
                '完成多边形
                If Not (marrXY(UBound(marrXY)).x = x And marrXY(UBound(marrXY)).y = y) _
                    And Not (marrXY(UBound(marrXY)).x = marrXY(1).x _
                    And marrXY(UBound(marrXY)).y = marrXY(1).y) Then
                    '最后两点相同则当作一点
                    '最后点与第一点相同则不处理
                    ReDim Preserve marrXY(1 To UBound(marrXY) + 1)
                    marrXY(UBound(marrXY)).x = x: marrXY(UBound(marrXY)).y = y
                End If
                Polygon picDraw.hDC, marrXY(1), UBound(marrXY)
                picDraw.Refresh '必须刷新
                
                '加入集合时用原始尺寸
                strXY = GetStrFromArr
                Call GetRect(strXY, X1, Y1, X2, Y2) '最小的范围
                mintKey = mintKey + 1
                mobjMapItems.Add 4, "", "", strXY, X1, Y1, X2, Y2, gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
                Call NewOper(mobjMapItems.Count, 1)
                
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            ElseIf UBound(marrXY) = 1 Then
                If marrXY(1).x = x And marrXY(1).y = y Then
                    '相同点取消作图
                    Call GetBufferAll '恢复原始图象
                    Call ResetDrawStyle
                    ReDim marrXY(0)
                    mlngTmpX = 0: mlngTmpY = 0
                    Exit Sub
                End If
                '多边形(第一段线的确认)
                MoveToEx picDraw.hDC, mlngOrgX, mlngOrgY, 0
                LineTo picDraw.hDC, x, y
                picDraw.Refresh
                ReDim Preserve marrXY(1 To 2): marrXY(2).x = x: marrXY(2).y = y
                mlngTmpX = 0: mlngTmpY = 0 '之后第一次作图不需要恢复
            End If
        End If
    ElseIf Button = 2 Then
        If tbrTool.Buttons("MLine").Value = tbrPressed Then
            If mblnDblClick And UBound(marrXY) >= 2 Then
                '取消本次作图
                Call GetBufferAll '恢复原始图象
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            End If
        ElseIf tbrTool.Buttons("MRect").Value = tbrPressed Then
            If mblnDblClick And UBound(marrXY) >= 2 Then
                '取消本次作图
                Call GetBufferAll '恢复原始图象
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            End If
        End If
    End If
End Sub

Private Sub picTxt_DblClick()
    Dim W As Long, H As Long, sngFont As Single
    
    '设置字体
    cdg.Flags = &H3 Or &H100 Or &H400 Or &H200 Or &H10000
    cdg.FontName = txt.FontName
    cdg.FontSize = txt.FontSize
    cdg.FontBold = txt.FontBold
    cdg.FontItalic = txt.FontItalic
    cdg.FontUnderline = txt.FontUnderline
    cdg.FontStrikethru = txt.FontStrikethru
    cdg.COLOR = txt.ForeColor
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowFont
    If Err.Number = 0 Then
        If cdg.FontSize > txt.FontSize Then
            Call GetTxtSize(cdg, txt.Text, W, H)
            If txt.Left + W + 10 <= picDraw.ScaleWidth And txt.Top + H + 6 <= picDraw.ScaleHeight Then
                sngFont = cdg.FontSize
            Else
                sngFont = txt.FontSize
            End If
            If sngFont < cdg.FontSize Then
                MsgBox "你设置的字体太大,文本无法在可见范围内完全显示,请调整文字位置或内容。", vbInformation, gstrSysName
            End If
        Else
            sngFont = cdg.FontSize
        End If
        txt.FontName = cdg.FontName
        
        txt.FontSize = sngFont
        txt.FontBold = cdg.FontBold
        txt.FontItalic = cdg.FontItalic
        txt.FontUnderline = cdg.FontUnderline
        txt.FontStrikethru = cdg.FontStrikethru
        txt.ForeColor = cdg.COLOR
                
        txt.Tag = txt.FontName & "," & txt.FontSize / msngScale & "," & txt.ForeColor & "," & _
            IIf(txt.FontBold, "1", "0") & IIf(txt.FontItalic, "1", "0") & _
            IIf(txt.FontUnderline, "1", "0") & IIf(txt.FontStrikethru, "1", "0")
        
        Call txt_Change
    End If
    txt.SetFocus
End Sub

Private Sub picTxt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngOrgX = x: mlngOrgY = y
End Sub

Private Sub picTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If txt.Left + x - mlngOrgX >= 0 And txt.Left + x - mlngOrgX + txt.Width <= picDraw.ScaleWidth Then
            picTxt.Left = picTxt.Left + x - mlngOrgX
            txt.Left = txt.Left + x - mlngOrgX
        End If
        If txt.Top + y - mlngOrgY >= 0 And txt.Top + y - mlngOrgY + txt.Height <= picDraw.ScaleHeight Then
            picTxt.Top = picTxt.Top + y - mlngOrgY
            txt.Top = txt.Top + y - mlngOrgY
        End If
        picDraw.Refresh
    End If
End Sub

Private Sub picTxt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txt.SetFocus
End Sub

Private Sub scrLR_Change()
    picDraw.Left = -scrLR.Value
End Sub

Private Sub scrLR_Scroll()
    scrLR_Change
End Sub

Private Sub scrUD_Change()
    picDraw.Top = -scrUD.Value
End Sub

Private Sub scrUD_Scroll()
    scrUD_Change
End Sub

Private Sub tbrStyle_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "FillColor"
            cdg.CancelError = True
            cdg.Flags = &H1
            cdg.COLOR = tbrStyle.Buttons("FillColor").Tag
            On Error Resume Next
            cdg.ShowColor
            If Err.Number = 0 Then
                LockWindowUpdate cbrStyle.hWnd
                
                tbrStyle.Buttons("FillColor").Image = 0
                tbrStyle.Buttons("LineColor").Image = 0
                img16.ListImages.Remove img16.ListImages.Count
                img16.ListImages.Remove img16.ListImages.Count
                
                picColor.BackColor = cdg.COLOR
                img16.ListImages.Add , , picColor.Image
                tbrStyle.Buttons("FillColor").Image = img16.ListImages.Count
                tbrStyle.Buttons("FillColor").Tag = cdg.COLOR
                
                picColor.BackColor = tbrStyle.Buttons("LineColor").Tag
                img16.ListImages.Add , , picColor.Image
                tbrStyle.Buttons("LineColor").Image = img16.ListImages.Count
                
                LockWindowUpdate 0
            End If
        Case "LineColor"
            cdg.CancelError = True
            cdg.Flags = &H1
            cdg.COLOR = tbrStyle.Buttons("LineColor").Tag
            On Error Resume Next
            cdg.ShowColor
            If Err.Number = 0 Then
                LockWindowUpdate cbrStyle.hWnd
                
                tbrStyle.Buttons("FillColor").Image = 0
                tbrStyle.Buttons("LineColor").Image = 0
                img16.ListImages.Remove img16.ListImages.Count
                img16.ListImages.Remove img16.ListImages.Count
                
                picColor.BackColor = tbrStyle.Buttons("FillColor").Tag
                img16.ListImages.Add , , picColor.Image
                tbrStyle.Buttons("FillColor").Image = img16.ListImages.Count
                
                picColor.BackColor = cdg.COLOR
                img16.ListImages.Add , , picColor.Image
                tbrStyle.Buttons("LineColor").Image = img16.ListImages.Count
                tbrStyle.Buttons("LineColor").Tag = cdg.COLOR
                
                LockWindowUpdate 0
            End If
        Case "LineDot", "LineDash", "LineDashDot", "LineDashDot2"
            '线型只能应用于宽度1
            tbrStyle.Buttons("Line1").Value = tbrPressed
        Case "Line2", "Line3", "Line4", "Line5"
            '非线宽1不能用线型
            tbrStyle.Buttons("LineAll").Value = tbrPressed
    End Select
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strXY As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    Dim sngStep As Single, objMapItem As MapItem, i As Integer

    '自动完成或取消
    If UBound(marrXY) >= 2 Then
        If mstrTool = "MLine" Then
            If Button.Key = "UnDo" Or UBound(marrXY) = 2 Then
                '取消画折线
                Call GetBufferAll
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            Else
                '完成画折线
                If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                    Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
                End If
                picDraw.Refresh
                
                strXY = GetStrFromArr
                Call GetRect(strXY, X1, Y1, X2, Y2) '最小的范围
                mintKey = mintKey + 1
                mobjMapItems.Add 2, "", "", strXY, X1, Y1, X2, Y2, gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
                Call NewOper(mobjMapItems.Count, 1)
                
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            End If
        ElseIf mstrTool = "MRect" Then
            If Button.Key = "UnDo" Or UBound(marrXY) = 2 Then
                '取消画多边形
                Call GetBufferAll
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            Else '完成画多边形
                If mlngTmpX <> 0 Or mlngTmpY <> 0 Then
                    Call GetBuffer(marrXY(UBound(marrXY)).x, marrXY(UBound(marrXY)).y, mlngTmpX, mlngTmpY, msngScale)
                End If
                Polygon picDraw.hDC, marrXY(1), UBound(marrXY)
                picDraw.Refresh
                
                strXY = GetStrFromArr
                Call GetRect(strXY, X1, Y1, X2, Y2) '最小的范围
                mintKey = mintKey + 1
                mobjMapItems.Add 4, "", "", strXY, X1, Y1, X2, Y2, gcurFillColor, gcurFillStyle, gcurPenColor, gcurPenStyle, gcurPenWidth, "_" & mintKey
                Call NewOper(mobjMapItems.Count, 1)
                
                Call ResetDrawStyle
                ReDim marrXY(0)
                mlngTmpX = 0: mlngTmpY = 0
            End If
        End If
    Else
        If txt.Visible And Button.Key <> "UnDo" And Button.Key <> "ReDo" Then Call FinishInput
        Select Case Button.Key
            Case "Move"
                Set picDraw.MouseIcon = imgCur.ListImages("Move").Picture
            Case "Text"
                Set picDraw.MouseIcon = imgCur.ListImages("Text").Picture
            Case "Line", "MLine", "Rect", "MRect", "Circle"
                Set picDraw.MouseIcon = imgCur.ListImages("Pen").Picture
            Case "Earse"
                Set picDraw.MouseIcon = imgCur.ListImages("Earse").Picture
            Case "ZoomIn"
                If msngScale >= 1 Then
                    sngStep = 0.5
                Else
                    sngStep = 0.25
                End If
                If msngScale + sngStep <= 5 Then
                    msngScale = msngScale + sngStep
                    Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
                End If
                Call Form_Resize
            Case "ZoomOut"
                If msngScale > 1 Then
                    sngStep = 0.5
                Else
                    sngStep = 0.25
                End If
                If msngScale - sngStep >= 0.25 Then
                    msngScale = msngScale - sngStep
                    Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
                End If
                Call Form_Resize
            Case "ZoomNone"
                If msngScale <> 1 Then
                    msngScale = 1
                    Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
                    Call Form_Resize
                End If
            Case "Clear"
                If mobjMapItems.Count = 0 Then Exit Sub
                
                For i = 1 To mobjMapItems.Count
                    Call NewOper(i, 3)
                Next
                
                Set mobjMapItems = New MapItems
                Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
            Case "UnDo"
                If txt.Visible Then
                    SendMessage txt.hWnd, WM_UNDO, 0, 0 '文本框内的取消和重做
                Else
                    If mintOper > 0 Then
                        Set objMapItem = mcolOper(mintOper)
                        If objMapItem.Oper = 1 Then '撤消增加
                            mobjMapItems.Remove "_" & Split(objMapItem.Key, "_")(1) '注意原始关键字
                            mintOper = mintOper - 1
                            
                            '文字再做一步
                            If mintOper > 0 Then
                                Set objMapItem = mcolOper(mintOper)
                                '删除跟STAMP 文字相关的椭圆
                                If objMapItem.Oper = 1 And objMapItem.类型 = 5 And objMapItem.点集 = "Stamp" Then
                                    mobjMapItems.Remove "_" & Split(objMapItem.Key, "_")(1) '注意原始关键字
                                    mintOper = mintOper - 1
                                    mintStampNo = mintStampNo - 1
                                End If
                                
                                If objMapItem.Oper = 2 Then '撤消删除
                                    With objMapItem
                                        mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                                    End With
                                    mintOper = mintOper - 1
                                End If
                                
                            End If
                        ElseIf objMapItem.Oper = 2 Then '撤消删除
                            With objMapItem
                                mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                            End With
                            mintOper = mintOper - 1
                        ElseIf objMapItem.Oper = 3 Then '撤消清除(连续删除)
                            Do While objMapItem.Oper = 3 And mintOper > 0
                                With objMapItem
                                    mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                                End With
                                mintOper = mintOper - 1
                                If mintOper > 0 Then Set objMapItem = mcolOper(mintOper)
                            Loop
                        End If
                        Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
                    End If
                    Call SetOperState
                End If
            Case "ReDo"
                If txt.Visible Then
                    SendMessage txt.hWnd, WM_UNDO, 0, 0 '文本框内的取消和重做
                Else
                    If mintOper < mcolOper.Count Then
                        Set objMapItem = mcolOper(mintOper + 1)
                        If objMapItem.Oper = 1 Then '重做增加
                            mintOper = mintOper + 1
                            With objMapItem
                                mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                            End With
                            
                            '重做 STAMP类型中跟文字关联的椭圆
                            If mintOper < mcolOper.Count Then
                                Set objMapItem = mcolOper(mintOper + 1)
                                If objMapItem.Oper = 1 And objMapItem.类型 = 0 And objMapItem.点集 = "Stamp" Then
                                    mintOper = mintOper + 1
                                    mintStampNo = mintStampNo + 1
                                    With objMapItem
                                        mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                                    End With
                                End If
                            End If
                        ElseIf objMapItem.Oper = 2 Then '重做删除
                            mintOper = mintOper + 1
                            mobjMapItems.Remove "_" & Split(objMapItem.Key, "_")(1)
                            
                            '文字再做一步
                            If mintOper < mcolOper.Count Then
                                Set objMapItem = mcolOper(mintOper + 1)
                                If objMapItem.Oper = 1 Then '重做增加
                                    mintOper = mintOper + 1
                                    With objMapItem
                                        mobjMapItems.Add .类型, .内容, .字体, .点集, .X1, .Y1, .X2, .Y2, .填充色, .填充方式, .线条色, .线型, .线宽, "_" & Split(.Key, "_")(1)
                                    End With
                                End If
                            End If
                        ElseIf objMapItem.Oper = 3 Then '重做清除(连续删除)
                            Do While objMapItem.Oper = 3 And mintOper < mcolOper.Count
                                mobjMapItems.Remove "_" & Split(objMapItem.Key, "_")(1)
                                mintOper = mintOper + 1
                                If mintOper + 1 <= mcolOper.Count Then Set objMapItem = mcolOper(mintOper + 1)
                            Loop
                        End If
                        Call ShowCaseMap(picDraw, mobjCaseMap, mobjMapItems, msngScale, picMap)
                    End If
                    Call SetOperState
                End If
            Case "Save"
                gblnOK = True
                Unload Me
            Case "Exit"
                Unload Me
        End Select
    End If
    If Button.Style = tbrButtonGroup Then mstrTool = Button.Key
End Sub

Private Sub SetDrawStyleFromFace(sngScale As Single)
'功能：根据界面状态设置当前的画笔的画刷
    Dim bytPenW As Byte
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '画笔
    If tbrStyle.Buttons("Line1").Value = tbrPressed Then
        bytPenW = 1
    ElseIf tbrStyle.Buttons("Line2").Value = tbrPressed Then
        bytPenW = 2
    ElseIf tbrStyle.Buttons("Line3").Value = tbrPressed Then
        bytPenW = 3
    ElseIf tbrStyle.Buttons("Line4").Value = tbrPressed Then
        bytPenW = 4
    ElseIf tbrStyle.Buttons("Line5").Value = tbrPressed Then
        bytPenW = 5
    End If
    gcurPenWidth = bytPenW '记录原始数据
    bytPenW = bytPenW * sngScale
    If bytPenW < 1 Then bytPenW = 1
    
    gcurPenColor = Val(tbrStyle.Buttons("LineColor").Tag)
    If tbrStyle.Buttons("LineAll").Value = tbrPressed Then
        gcurPenStyle = PS_SOLID
        lngPen = CreatePen(PS_SOLID, bytPenW, Val(tbrStyle.Buttons("LineColor").Tag))
    ElseIf tbrStyle.Buttons("LineDot").Value = tbrPressed Then
        gcurPenStyle = PS_DOT
        lngPen = CreatePen(PS_DOT, bytPenW, Val(tbrStyle.Buttons("LineColor").Tag))
    ElseIf tbrStyle.Buttons("LineDash").Value = tbrPressed Then
        gcurPenStyle = PS_DASH
        lngPen = CreatePen(PS_DASH, bytPenW, Val(tbrStyle.Buttons("LineColor").Tag))
    ElseIf tbrStyle.Buttons("LineDashDot").Value = tbrPressed Then
        gcurPenStyle = PS_DASHDOT
        lngPen = CreatePen(PS_DASHDOT, bytPenW, Val(tbrStyle.Buttons("LineColor").Tag))
    ElseIf tbrStyle.Buttons("LineDashDot2").Value = tbrPressed Then
        gcurPenStyle = PS_DASHDOTDOT
        lngPen = CreatePen(PS_DASHDOTDOT, bytPenW, Val(tbrStyle.Buttons("LineColor").Tag))
    End If
    glngPen = SelectObject(picDraw.hDC, lngPen)
    
    '画刷
    vBrush.lbColor = Val(tbrStyle.Buttons("FillColor").Tag)
    gcurFillColor = vBrush.lbColor
    If tbrStyle.Buttons("FillNone").Value = tbrPressed Then
        vBrush.lbStyle = BS_NULL
        gcurFillStyle = -1
    ElseIf tbrStyle.Buttons("FillAll").Value = tbrPressed Then
        vBrush.lbStyle = BS_SOLID
        gcurFillStyle = -2
    Else
        vBrush.lbStyle = BS_HATCHED
        If tbrStyle.Buttons("FillHsc").Value = tbrPressed Then
            vBrush.lbHatch = HS_HORIZONTAL '====
        ElseIf tbrStyle.Buttons("FillVsc").Value = tbrPressed Then
            vBrush.lbHatch = HS_VERTICAL '||||
        ElseIf tbrStyle.Buttons("FillHV").Value = tbrPressed Then
            vBrush.lbHatch = HS_CROSS '++++
        ElseIf tbrStyle.Buttons("FillL").Value = tbrPressed Then
            vBrush.lbHatch = HS_FDIAGONAL '\\\\
        ElseIf tbrStyle.Buttons("FillR").Value = tbrPressed Then
            vBrush.lbHatch = HS_BDIAGONAL '////
        ElseIf tbrStyle.Buttons("FillLR").Value = tbrPressed Then
            vBrush.lbHatch = HS_DIAGCROSS 'XXXX
        End If
        gcurFillStyle = vBrush.lbHatch
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(picDraw.hDC, lngBrush)
End Sub

Private Sub ForceSquare(ByVal X1 As Long, ByVal Y1 As Long, X2 As Long, Y2 As Long)
'功能：将指定矩形坐标强行调整成正方形
'返回：x2,y2=新的矩形结束点
    If Abs(Y2 - Y1) > Abs(X2 - X1) Then
        If X2 < X1 Then
            X2 = X1 - Abs(Y2 - Y1)
        Else
            X2 = X1 + Abs(Y2 - Y1)
        End If
        Y2 = Y2
    Else
        If Y2 < Y1 Then
            Y2 = Y1 - Abs(X2 - X1)
        Else
            Y2 = Y1 + Abs(X2 - X1)
        End If
        X2 = X2
    End If
End Sub

Private Sub SetBuffer(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, sngScale As Single)
'功能：缓冲当前作图指定区域
'说明：1.用作缓冲的PictureBox.AutoRedraw必须设置为True,但不用再调用Refresh方法
'      2.用作缓冲的PictureBox的尺寸必须大于要备份的区域尺寸
    Dim x As Long, y As Long, W As Long, H As Long
    If X2 < X1 Then
        x = X2 - 3 * sngScale
        W = X1 - X2 + 7 * sngScale '1 + 3 * 2
    Else
        x = X1 - 3 * sngScale
        W = X2 - X1 + 7 * sngScale
    End If
    If Y2 < Y1 Then
        y = Y2 - 3 * sngScale
        H = Y1 - Y2 + 7 * sngScale
    Else
        y = Y1 - 3 * sngScale
        H = Y2 - Y1 + 7 * sngScale
    End If
    StretchBlt picBuf.hDC, 0, 0, W, H, picDraw.hDC, x, y, W, H, SRCCOPY
End Sub

Private Sub GetBuffer(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, sngScale As Single)
'功能：从缓冲中恢复指定作图区域
    Dim x As Long, y As Long, W As Long, H As Long
    If X2 < X1 Then
        x = X2 - 3 * sngScale
        W = X1 - X2 + 7 * sngScale
    Else
        x = X1 - 3 * sngScale
        W = X2 - X1 + 7 * sngScale
    End If
    If Y2 < Y1 Then
        y = Y2 - 3 * sngScale
        H = Y1 - Y2 + 7 * sngScale
    Else
        y = Y1 - 3 * sngScale
        H = Y2 - Y1 + 7 * sngScale
    End If
    StretchBlt picDraw.hDC, x, y, W, H, picBuf.hDC, 0, 0, W, H, SRCCOPY
End Sub

Private Sub SetBufferAll()
'功能：将当前整个图象备份到缓冲区
    BitBlt picOrig.hDC, 0, 0, picDraw.ScaleWidth, picDraw.ScaleHeight, picDraw.hDC, 0, 0, SRCCOPY
End Sub

Private Sub GetBufferAll()
'功能：从缓冲区取回整个备份图象
    BitBlt picDraw.hDC, 0, 0, picDraw.ScaleWidth, picDraw.ScaleHeight, picOrig.hDC, 0, 0, SRCCOPY
    picDraw.Refresh
End Sub

Private Function GetStrFromArr() As String
    Dim i As Integer, str As String
    
    If UBound(marrXY) = 0 Then Exit Function
    For i = 1 To UBound(marrXY)
        str = str & ";" & marrXY(i).x / msngScale & "," & marrXY(i).y / msngScale
    Next
    GetStrFromArr = Mid(str, 2)
End Function

Private Function InLine(ByVal x As Long, ByVal y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal PenW As Byte = 1) As Boolean
'功能：判断一个点(x,y)是否在直线(x1,y1-x2,y2)上
    Dim A As Double, B As Double, C As Double, k As Double
        
    k = IIf(PenW / 2 * msngScale < 2, 2, PenW / 2 * msngScale)  '误差点范围
    
    If X2 < X1 Then
        InLine = x >= X2 And x <= X1
    ElseIf X2 > X1 Then
        InLine = x >= X1 And x <= X2
    ElseIf X2 = X1 Then
        InLine = Abs(x - X1) <= k
    End If
    If Y2 < Y1 Then
        InLine = InLine And y >= Y2 And y <= Y1
    ElseIf Y2 > Y1 Then
        InLine = InLine And y >= Y1 And y <= Y2
    ElseIf Y1 = Y2 Then
        InLine = InLine And Abs(y - Y1) <= k
    End If
    
    If InLine Then
        A = X2 - X1
        B = Y2 - Y1
        C = Y1 * A - X1 * B
        'ay-bx=c;y=(c+bx)/a;x=(ay-c)/b
        If A <> 0 And B <> 0 Then
            InLine = Abs(y - (C + B * x) / A) <= k Or Abs(x - (A * y - C) / B) <= k
        ElseIf A <> 0 Then
            InLine = Abs(y - (C + B * x) / A) <= k
        ElseIf B <> 0 Then
            InLine = Abs(x - (A * y - C) / B) <= k
        Else '两点相同的点线
            InLine = Abs(x - X1) <= k Or Abs(y - Y1) <= k
        End If
    End If
End Function

Private Function InEllipse(ByVal x As Long, ByVal y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal PenW As Byte = 1) As Boolean
'功能：判断一点是否在一个椭圆上
    Dim lngErr As Long, lngTmp As Long
    Dim Ox As Long, Oy As Long
    Dim A As Double, B As Double
    
    '误差点范围
    lngErr = IIf(PenW * msngScale < 2, 2, PenW * msngScale)
    
    '转换范围(成左上至右下)
    If X2 < X1 Then lngTmp = X1: X1 = X2: X2 = lngTmp
    If Y2 < Y1 Then lngTmp = Y1: Y1 = Y2: Y2 = lngTmp
    
    '圆心位置,即与原点的距离
    Ox = X1 + (X2 - X1 + 1) / 2 - 1
    Oy = Y1 + (Y2 - Y1 + 1) / 2 - 1
    
    '将椭圆移到以原点为中心
    x = Abs(x - Ox): X1 = X1 - Ox: X2 = X2 - Ox
    y = Abs(y - Oy): Y1 = Y1 - Oy: Y2 = Y2 - Oy
    
    '使用椭圆的标准方程计算
    A = X2 ^ 2 'a^2
    B = Y2 ^ 2 'b^2
    If B <> 0 And A <> 0 Then
        If A - A * y ^ 2 / B >= 0 Then
            InEllipse = Abs(x - Sqr(A - A * y ^ 2 / B)) <= lngErr
        ElseIf B - B * x ^ 2 / A >= 0 Then
            InEllipse = Abs(y - Sqr(B - B * x ^ 2 / A)) <= lngErr
        End If
    End If
End Function

Private Function inText(ByVal x As Long, ByVal y As Long) As Integer
'功能：判断一个点下面是否有一段文字。
    Dim i As Integer
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    
    If mobjMapItems.Count = 0 Then Exit Function
    For i = 1 To mobjMapItems.Count
        If mobjMapItems(i).类型 = 0 Then
            With mobjMapItems(i)
                If .X2 < .X1 Then
                    X1 = .X2: X2 = .X1
                Else
                    X1 = .X1: X2 = .X2
                End If
                If .Y2 < .Y1 Then
                    Y1 = .Y2: Y2 = .Y1
                Else
                    Y1 = .Y1: Y2 = .Y2
                End If
                If x >= X1 * msngScale And x <= X2 * msngScale And y >= Y1 * msngScale And y <= Y2 * msngScale Then
                    inText = i: Exit Function
                End If
            End With
        End If
    Next
End Function

Private Sub ShowCaseMap(picShow As Object, objCaseMap As StdPicture, objMapItems As MapItems, sngScale As Single, objOrig As Object)
'功能：显示病历标记图内容
'参数：picShow=显示的目标对象
'      objMapItems=病历中当前项目的标记图内容
'      sngSclae=显示比例
'      objOrig=辅助的PictureBox控件(无边框,AutoSize,AutoRedraw)
    Dim lngW As Long, lngH As Long
    Dim i As Integer, j As Integer
    Dim arrTmp() As String, arrXY() As POINTAPI
        
    Screen.MousePointer = 11
    LockWindowUpdate picShow.hWnd
    
    picShow.Cls
    
    '尺寸及背景图
    If objOrig.Picture.Handle = 0 Then Set objOrig.Picture = objCaseMap
    
    lngW = objOrig.Width
    lngH = objOrig.Height
    
    picShow.Width = lngW * sngScale + 2
    picShow.Height = lngH * sngScale + 2
            
    StretchBlt picShow.hDC, 0, 0, picShow.ScaleWidth, picShow.ScaleHeight, objOrig.hDC, 0, 0, objOrig.Width, objOrig.Height, SRCCOPY
            
    '具体标记元素
    For i = 1 To objMapItems.Count
        With objMapItems(i)
            If .类型 <> 0 Then
                Call SetDrawStyleFromValue(picShow.hDC, .线条色, .线型, .线宽 * sngScale, .填充色, .填充方式)
            End If
            Select Case .类型
                Case 0 '文本
                    Call TextOut(picShow, .内容, .X1, .Y1, .X2, .Y2, .字体, sngScale)
                Case 1 '线条
                    MoveToEx picShow.hDC, .X1 * sngScale, .Y1 * sngScale, 0
                    LineTo picShow.hDC, .X2 * sngScale, .Y2 * sngScale
                Case 2 '折线
                    arrTmp = Split(.点集, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale
                    Next
                    Polyline picShow.hDC, arrXY(0), UBound(arrXY) + 1
                Case 3 '矩形
                    Rectangle picShow.hDC, .X1 * sngScale, .Y1 * sngScale, .X2 * sngScale, .Y2 * sngScale
                Case 4 '多边形
                    arrTmp = Split(.点集, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale
                    Next
                    Polygon picShow.hDC, arrXY(0), UBound(arrXY) + 1
                Case 5 '圆
                    Ellipse picShow.hDC, .X1 * sngScale, .Y1 * sngScale, .X2 * sngScale, .Y2 * sngScale
            End Select
        End With
    Next
    picShow.Refresh
    
    picBuf.Cls
    picBuf.Width = picShow.Width
    picBuf.Height = picShow.Height
    
    picOrig.Cls
    picOrig.Width = picShow.Width
    picOrig.Height = picShow.Height
    
    Call ResetDrawStyle
    
    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Private Sub DrawItemState(intItem As Integer, blnSel As Boolean)
'功能：以XOR或Copy方式画一个标注元素,表示是否选中
'参数：intItem=元素索引
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer
    
    If intItem = 0 Then Exit Sub
    If blnSel Then
        picDraw.DrawMode = vbInvert
    Else
        picDraw.DrawMode = vbCopyPen
    End If
    
    With mobjMapItems(intItem)
        If .类型 <> 0 Then Call SetDrawStyleFromValue(picDraw, .线条色, .线型, .线宽 * msngScale, .填充色, .填充方式)
        Select Case .类型
            Case 0 '文本
                
            Case 1 '线条
                If blnSel Then
                    Call SetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                    MoveToEx picDraw.hDC, .X1 * msngScale, .Y1 * msngScale, 0
                    LineTo picDraw.hDC, .X2 * msngScale, .Y2 * msngScale
                Else
                    Call GetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                End If
            Case 2 '折线
                If blnSel Then
                    Call SetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                    arrTmp = Split(.点集, ";")
                    For i = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(i)
                        arrXY(i).x = CLng(Split(arrTmp(i), ",")(0)) * msngScale
                        arrXY(i).y = CLng(Split(arrTmp(i), ",")(1)) * msngScale
                    Next
                    Polyline picDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Else
                    Call GetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                End If
            Case 3 '矩形
                If blnSel Then
                    Call SetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                    Rectangle picDraw.hDC, .X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale
                Else
                    Call GetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                End If
            Case 4 '多边形
                If blnSel Then
                    Call SetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                    arrTmp = Split(.点集, ";")
                    For i = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(i)
                        arrXY(i).x = CLng(Split(arrTmp(i), ",")(0)) * msngScale
                        arrXY(i).y = CLng(Split(arrTmp(i), ",")(1)) * msngScale
                    Next
                    Polygon picDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Else
                    Call GetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                End If
            Case 5 '(椭)圆
                If blnSel Then
                    Call SetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                    Ellipse picDraw.hDC, .X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale
                Else
                    Call GetBuffer(.X1 * msngScale, .Y1 * msngScale, .X2 * msngScale, .Y2 * msngScale, msngScale)
                End If
        End Select
    End With
    picDraw.Refresh
    If blnSel Then picDraw.DrawMode = vbCopyPen
End Sub

Private Sub GetRect(ByVal strXY As String, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
'功能：返回一个能全部包含一个多边形的最小矩形范围
'参数：strXY=点集
    Dim i As Integer
    Dim arrTmp() As String, arrXY() As String
    
    arrTmp = Split(strXY, ";")
    For i = 0 To UBound(arrTmp)
        arrXY = Split(arrTmp(i), ",")
        If i = 0 Then
            X1 = arrXY(0): X2 = X1
            Y1 = arrXY(1): Y2 = Y1
        Else
            If arrXY(0) < X1 Then
                X1 = arrXY(0)
            ElseIf arrXY(0) > X2 Then
                X2 = arrXY(0)
            End If
            If arrXY(1) < Y1 Then
                Y1 = arrXY(1)
            ElseIf arrXY(1) > Y2 Then
                Y2 = arrXY(1)
            End If
        End If
    Next
End Sub

Private Sub txt_Change()
    Dim W As Long, h2 As Long
    Dim lngLines As Long
    
    Call GetTxtSize(txt, txt.Text, W, , h2)
    
    If txt.Left + W + 10 <= picDraw.ScaleWidth Then
        txt.Width = W + 10
        picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
    End If
    
    lngLines = SendMessage(txt.hWnd, EM_GETLINECOUNT, 0, 0)
    txt.Height = lngLines * h2 + 6
    picTxt.Top = txt.Top - picTxt.Height / 2
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim h2 As Long, lngLines As Long
    
    If InStr("'%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub '非法
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0: Beep: Exit Sub  '超长
    
    If KeyAscii >= 32 Or KeyAscii = 13 Or KeyAscii < 0 Then
        txtTmp.FontSize = txt.FontSize
        txtTmp.FontName = txt.FontName
        txtTmp.FontBold = txt.FontBold
        txtTmp.FontItalic = txt.FontItalic
        txtTmp.FontUnderline = txt.FontUnderline
        txtTmp.FontStrikethru = txt.FontStrikethru
        txtTmp.Width = txt.Width
        txtTmp.Text = Left(txt.Text, txt.SelStart) & IIf(KeyAscii = 13, vbCrLf, Chr(KeyAscii)) & Mid(txt.Text, txt.SelStart + txt.SelLength + 1)
        lngLines = SendMessage(txtTmp.hWnd, EM_GETLINECOUNT, 0, 0)
        Call GetTxtSize(txt, "A", , , h2)
        If txt.Top + lngLines * h2 + 6 > picDraw.ScaleHeight Then KeyAscii = 0: Beep
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    If txt.Left + txt.Width > picDraw.ScaleWidth Or txt.Top + txt.Height > picDraw.Height Then
        Cancel = True
        MsgBox "文本内容无法在可见范围内完全显示,请调整文本位置或内容！", vbInformation, gstrSysName
    End If
End Sub

Private Sub FinishInput()
'功能：完成当前文字输入
    If txt.Visible Then
        '从输入状态转为确定输入并退出
        If Trim(Replace(txt.Text, vbCrLf, "")) <> "" Then
            '加入文字项
            mintKey = mintKey + 1
            mobjMapItems.Add 0, txt.Text, txt.FontName & "," & txt.FontSize / msngScale & "," & txt.ForeColor & "," & _
                IIf(txt.FontBold, "1", "0") & IIf(txt.FontItalic, "1", "0") & IIf(txt.FontUnderline, "1", "0") & IIf(txt.FontStrikethru, "1", "0"), _
                "", txt.Left / msngScale, txt.Top / msngScale, (txt.Left + txt.Width) / msngScale, (txt.Top + txt.Height) / msngScale, 0, 0, 0, 0, 0, "_" & mintKey
            
            Call NewOper(mobjMapItems.Count, 1)
            
            With mobjMapItems(mobjMapItems.Count)
                TextOut picDraw, .内容, .X1, .Y1, .X2, .Y2, .字体, msngScale
            End With
        End If
        txt.Text = ""
        txt.Visible = False
        picTxt.Visible = False
        
        Call SetOperState
    End If
End Sub

Private Sub NewOper(intItem As Integer, intOper As Integer)
'功能：新记录一个操作
'参数：intItem=项目索引,intOper=1-增加,2-删除,3-清除
'说明：如果是增加,应该在增加之后记录,如果是删除,应该在删除之前记录
    Dim i As Integer
    Dim objMapItem As MapItem
    
    Set objMapItem = New MapItem
    
    '复制项目内容
    With mobjMapItems(intItem)
        objMapItem.类型 = .类型
        objMapItem.内容 = .内容
        objMapItem.字体 = .字体
        objMapItem.点集 = .点集
        objMapItem.X1 = .X1: objMapItem.Y1 = .Y1
        objMapItem.X2 = .X2: objMapItem.Y2 = .Y2
        objMapItem.填充方式 = .填充方式
        objMapItem.填充色 = .填充色
        objMapItem.线型 = .线型
        objMapItem.线宽 = .线宽
        objMapItem.线条色 = .线条色
        objMapItem.Key = intOper & .Key '同一个对象可能因为两个不同操作加入
    End With
    objMapItem.Oper = intOper
    
    '操作后,不能重做到以前撤消的操作
    If mcolOper.Count > mintOper Then
        For i = mcolOper.Count To mintOper + 1 Step -1
            mcolOper.Remove i
        Next
    End If
        
    '压入堆栈
    mintOper = mintOper + 1
    mcolOper.Add objMapItem, objMapItem.Key
    
    Set objMapItem = Nothing
    
    Call SetOperState
End Sub

Private Sub SetOperState()
'功能：根据操作堆栈,设置当前"撤消","恢复"功能状态
    tbrTool.Buttons("UnDo").Enabled = mintOper > 0 Or txt.Visible
    tbrTool.Buttons("ReDo").Enabled = mintOper < mcolOper.Count Or txt.Visible
End Sub
