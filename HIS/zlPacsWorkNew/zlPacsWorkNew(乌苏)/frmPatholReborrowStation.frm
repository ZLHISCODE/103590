VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholReborrowStation 
   Caption         =   "病理借还工作站"
   ClientHeight    =   9555
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13500
   Icon            =   "frmPatholReborrowStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   13500
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   7980
      Left            =   4455
      TabIndex        =   1
      Top             =   1080
      Width           =   100
      _ExtentX        =   185
      _ExtentY        =   14076
      BackColor       =   -2147483633
      SplitWidth      =   100
      SplitLevel      =   3
      SyncParentHeight=   0   'False
      AllowPaintOtherSpliter=   -1  'True
      Control1Name    =   "Picture1"
      Control2Name    =   "Picture2"
   End
   Begin VB.PictureBox picTimeOut 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5640
      Picture         =   "frmPatholReborrowStation.frx":179A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   4440
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgMenus 
      Left            =   3600
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":1ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":2180
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":2502
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":2854
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":2EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":324A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":359C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":38EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":3C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":3F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":42E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":4636
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":4988
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":4CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":502C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":537E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":56D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":5A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":5D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":60C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":6418
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":676A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":6ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":6E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":7160
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":74B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":7804
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2760
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":7B56
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":8830
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":950A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":A1E4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":AEBE
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":BB98
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":C872
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":D54C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":E226
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":EF00
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholReborrowStation.frx":FBDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7980
      Left            =   0
      ScaleHeight     =   7980
      ScaleWidth      =   4455
      TabIndex        =   5
      Top             =   1080
      Width           =   4455
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   120
         ScaleHeight     =   7575
         ScaleWidth      =   4335
         TabIndex        =   6
         Top             =   360
         Width           =   4335
         Begin zl9PacsControl.ucSplitter ucSplitter2 
            Height          =   100
            Left            =   0
            TabIndex        =   7
            Top             =   3930
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   185
            BackColor       =   -2147483633
            MousePointer    =   7
            SplitWidth      =   100
            SplitType       =   0
            SplitLevel      =   3
            Control1Name    =   "ufgBorrow"
            Control2Name    =   "rtbDetail"
         End
         Begin RichTextLib.RichTextBox rtbDetail 
            Height          =   3545
            Left            =   0
            TabIndex        =   8
            Top             =   4030
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   6244
            _Version        =   393217
            BackColor       =   16761024
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmPatholReborrowStation.frx":10C2C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin zl9PACSWork.ucFlexGrid ufgBorrow 
            Height          =   3930
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   6932
            GridRows        =   201
            BackColor       =   12648447
            IsEnterNextCell =   0   'False
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   0
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   661
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   7980
      Left            =   4555
      ScaleHeight     =   7980
      ScaleWidth      =   8940
      TabIndex        =   2
      Top             =   1080
      Width           =   8945
      Begin VB.Frame framMaterialDetail 
         Height          =   7695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8655
         Begin zl9PACSWork.ucFlexGrid ufgMaterialDetail 
            Height          =   4455
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7858
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin zl9PACSWork.ucFlexGrid ufgBackHistory 
            Height          =   2175
            Left            =   120
            TabIndex        =   13
            Top             =   5400
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3836
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label labBackHistory 
            Caption         =   "归还历史："
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   5160
            Width           =   975
         End
         Begin VB.Label labBorrowDetail 
            Caption         =   "借阅材料："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   9195
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPatholReborrowStation.frx":10CC9
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "未归还数量："
            TextSave        =   "未归还数量："
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "超期未归还数量："
            TextSave        =   "超期未归还数量："
            Key             =   "sb_NoEnterWaxStone"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
            Text            =   "部分归还数量："
            TextSave        =   "部分归还数量："
            Key             =   "sb_NoEnterSlices"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "已遗失借阅数量："
            TextSave        =   "已遗失借阅数量："
            Key             =   "sb_NoEnterSpeEx"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3177
            MinWidth        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "回执预览"
            Key             =   "tbn_PreviewBorrow"
            Object.Tag             =   "回执预览"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "回执打印"
            Key             =   "tbn_PrintBorrow"
            Object.Tag             =   "回执打印"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增借阅"
            Key             =   "tbn_NewLend"
            Object.Tag             =   "新增借阅"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除借阅"
            Key             =   "tbn_DelLend"
            Object.Tag             =   "删除借阅"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "更新借阅"
            Key             =   "tbn_UpdateLend"
            Object.Tag             =   "更新借阅"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询借阅"
            Key             =   "tbn_QueryLend"
            Object.Tag             =   "查询借阅"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "确认借阅"
            Key             =   "tbn_SureLend"
            Object.Tag             =   "确认借阅"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "撤销借阅"
            Key             =   "tbn_CancelLend"
            Object.Tag             =   "撤销借阅"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "借阅归还"
            Key             =   "tbn_ReturnLend"
            Object.Tag             =   "借阅归还"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            Key             =   "tbn_Help"
            Object.Tag             =   "帮助"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "tbn_Exit"
            Object.Tag             =   "退出"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnu_Flie 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnu_ParameterConfig 
         Caption         =   "参数设置(&C)"
      End
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "打印设置(&N)"
      End
      Begin VB.Menu mnu_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Preview 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportExcel 
         Caption         =   "输出到Execl(&E)"
      End
      Begin VB.Menu mnu_Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnu_NewBorrow 
         Caption         =   "新增借阅(&N)"
      End
      Begin VB.Menu mnu_DelBorrow 
         Caption         =   "删除借阅(&D)"
      End
      Begin VB.Menu mnu_UpdateBorrow 
         Caption         =   "更新借阅(&U)"
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SureBorrow 
         Caption         =   "确认借阅(&S)"
      End
      Begin VB.Menu mnu_CancelBorrow 
         Caption         =   "撤销借阅(&C)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_QueryBorrow 
         Caption         =   "借阅查询(&Q)"
      End
      Begin VB.Menu mnu_Split6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ReturnBorrow 
         Caption         =   "借阅归还(&R)"
      End
      Begin VB.Menu mnu_Split9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PreviewBorrow 
         Caption         =   "回执预览(&V)"
      End
      Begin VB.Menu mnu_PrintBorrow 
         Caption         =   "回执打印(&P)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnu_ToolBar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnu_StandardButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_WordLabel 
            Caption         =   "文本标签(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_StateBar 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Split7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Font 
         Caption         =   "字体(&F)"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "工具(&T)"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnu_HelpMan 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnu_Web 
         Caption         =   "WEB上的中联(&W)"
         Begin VB.Menu mnu_HomePage 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnu_bbs 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnu_Send 
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnu_Split8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "关于...(&A)"
      End
   End
End
Attribute VB_Name = "frmPatholReborrowStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False

'确认状态
Private Const BorrowSureState_NoSure As String = "未确认"
Private Const BorrowSureState_Sure As String = "已确认"

'归还状态
Private Const BorrowReturnState_Return As String = "已归还" '未归还，部分归还，遗失处理

'为菜单设置相应的图形
Private Const MF_BITMAP = &H400&


Private mstrPrivs As String
Private mlngDefaultQueryDays As Long
Private mstrBorrowReportName As String
Private mblnIsAutoPrint As Boolean
Private mblnMoved As Boolean

Private mdtServicesTime As Date

Private WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long '取得窗口的菜单句柄,hwnd是窗口的句柄
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal npos As Long) As Long '取得子菜单句柄，nPos是菜单的位置
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal npos As Long, ByVal wFlags As Long, ByVal hBitUnchecked As Long, ByVal hBitChecked As Long) As Long




Private Sub AdjustLayOut()
    
    Picture1.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    Picture1.Height = Me.ScaleHeight - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
    
    Call ucSplitter1.RePaint
End Sub


Private Sub LoadParameterConfig()
'载入相关参数配置
    mlngDefaultQueryDays = zlDatabase.GetPara("借阅默认查询天数", glngSys, G_LNG_PATHOLBORROW_NUM, "100")
    mstrBorrowReportName = zlDatabase.GetPara("借阅回执报表名称", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    mblnIsAutoPrint = zlDatabase.GetPara("借阅确认后自动打印回执", glngSys, G_LNG_PATHOLBORROW_NUM, 1)
End Sub



Private Sub ConfigPopedomFace()
'更加权限配置界面，如果不具备权限时，则隐藏对应功能按钮
    Dim i As Long
    
    mnu_ParameterConfig.Visible = CheckPopedom(mstrPrivs, "参数设置")
    
    mnu_CancelBorrow.Visible = CheckPopedom(mstrPrivs, "撤销借阅")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_CancelLend") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "撤销借阅")
        End If
    Next i
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
'    #If DebugState = True Then
'        Call InitDebugObject(1294, Me, "zlhis", "HIS")
'    #End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitTabs
    Call InitMenuIcoConfig
    
    Call InitBorrowList
    Call InitBorrowDetailList
    Call InitBackHistoryList
    
    Call LoadParameterConfig


    mstrPrivs = gstrPrivs
    
    Call ConfigPopedomFace
    
    Set zlReport = New zl9Report.clsReport
    
    curDate = zlDatabase.Currentdate
    
    Call QueryBorrowData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
    Call RefreshStateInf
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub RefreshStateInf()
'刷新材料遗失数量
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select " & _
            " sum(case when 归还状态=0 then 1 else 0 end)  as 未归还数量, " & _
            " sum(case when (借阅时间+借阅天数<sysdate and 归还状态=0) then 1 else 0 end)  as 超期未归还数量, " & _
            " sum(case when 归还状态=2 then 1 else 0 end)as 部分归还数量, " & _
            " sum(case when 归还状态=3 then 1 else 0 end) as 遗失数量 " & _
            " From 病理借阅信息 "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If rsData.RecordCount > 0 Then
        stbThis.Panels(2).Text = "未归还数量：" & Nvl(rsData!未归还数量)
        stbThis.Panels(3).Text = "超期未归还数量：" & Nvl(rsData!超期未归还数量)
        stbThis.Panels(4).Text = "部分归还数量：" & Nvl(rsData!部分归还数量)
        stbThis.Panels(5).Text = "已遗失借阅数量：" & Nvl(rsData!遗失数量)
    End If
End Sub


Private Sub InitMenuIcoConfig()
'初始化菜单图标显示
On Error Resume Next
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim hSubSubMenu As Long
    
    hMenu = GetMenu(Me.hWnd)
    
    '设置第一项菜单(文件)
    hSubMenu = GetSubMenu(hMenu, 0) '取得第一项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(28).Picture, imgMenus.ListImages(28).Picture) '参数设置
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(3).Picture, imgMenus.ListImages(3).Picture) '打印设置
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '打印预览
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '打印
    Call SetMenuItemBitmaps(hSubMenu, 6, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '导出Excel
    Call SetMenuItemBitmaps(hSubMenu, 8, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '退出
    

    '设置第二项菜单（编辑）
    hSubMenu = GetSubMenu(hMenu, 1) '取得第二项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '新增借阅
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(8).Picture, imgMenus.ListImages(8).Picture) '删除借阅
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(9).Picture, imgMenus.ListImages(9).Picture) '更新借阅
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(11).Picture, imgMenus.ListImages(11).Picture) '确认借阅
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '撤销借阅
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(12).Picture, imgMenus.ListImages(12).Picture) '查询借阅
    Call SetMenuItemBitmaps(hSubMenu, 9, MF_BITMAP, imgMenus.ListImages(29).Picture, imgMenus.ListImages(29).Picture) '借阅归还
    Call SetMenuItemBitmaps(hSubMenu, 11, MF_BITMAP, imgMenus.ListImages(1).Picture, imgMenus.ListImages(1).Picture) '回执预览
    Call SetMenuItemBitmaps(hSubMenu, 12, MF_BITMAP, imgMenus.ListImages(2).Picture, imgMenus.ListImages(2).Picture) '回执打印
    
    
    '设置第二项菜单（查看）
    hSubMenu = GetSubMenu(hMenu, 2) '取得第二项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(27).Picture, imgMenus.ListImages(27).Picture) '工具栏
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(22).Picture, imgMenus.ListImages(21).Picture) '状态栏
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(23).Picture, imgMenus.ListImages(23).Picture) '字体
    
        hSubSubMenu = GetSubMenu(hSubMenu, 0)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(26).Picture, imgMenus.ListImages(20).Picture) '标准按钮
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(25).Picture, imgMenus.ListImages(24).Picture) '文本标签
    
    
    
    '设置第五项菜单（帮助）
    hSubMenu = GetSubMenu(hMenu, 3) '取得第五项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(13).Picture, imgMenus.ListImages(13).Picture) '帮助主题
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(14).Picture) 'web中联
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(15).Picture, imgMenus.ListImages(15).Picture) '关
    
        hSubSubMenu = GetSubMenu(hSubMenu, 1)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(13).Picture) '帮助主题
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(16).Picture, imgMenus.ListImages(16).Picture) '中联论坛
        Call SetMenuItemBitmaps(hSubSubMenu, 2, MF_BITMAP, imgMenus.ListImages(17).Picture, imgMenus.ListImages(17).Picture) '发送反馈
    
    err.Clear

End Sub



Private Sub ReadBorrowInf(ByVal lngArchivesRowIndex As Long)
'读取借阅信息
    Dim strInf As String
    If lngArchivesRowIndex <= 0 Then Exit Sub
    
    strInf = "借 阅 号：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_借阅号) & vbCrLf
    strInf = strInf & "借 阅 人：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_借阅人) & vbCrLf
    strInf = strInf & "证件类型：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_证件类型) & vbCrLf
    strInf = strInf & "证件号码：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_证件号码) & vbCrLf
    strInf = strInf & "联系电话：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_联系电话) & vbCrLf
    strInf = strInf & "联系地址：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_联系地址) & vbCrLf
    strInf = strInf & "借阅押金：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_押金) & vbCrLf
    strInf = strInf & "借阅日期：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_借阅日期) & vbCrLf
    strInf = strInf & "借阅天数：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_借阅天数) & vbCrLf
    strInf = strInf & "归还日期：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_归还日期) & vbCrLf
    strInf = strInf & "借阅原因：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_借阅原因) & vbCrLf
    strInf = strInf & "归还状态：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_归还状态) & vbCrLf
    strInf = strInf & "备注说明：" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_备注)
    
    rtbDetail.Text = strInf
End Sub



Private Sub QueryBorrowData(ByVal dtStartDate As Date, ByVal dtEndDate As Date, _
    Optional ByVal strBorrowId As String = "", Optional ByVal strCardNo As String = "", _
    Optional ByVal strBorrowName As String = "")
'查询借阅记录数据
    Dim strSql As String
    Dim strFilter As String
    
    
    mdtServicesTime = zlDatabase.Currentdate
    
    strFilter = ""
    
    strFilter = " 借阅时间 between [1] and [2]"
    
    If Trim(strCardNo) <> "" Then
        strFilter = strFilter & " and upper(证件号码)=upper([4])"
    End If
    
    If Trim(strBorrowName) <> "" Then
        strFilter = strFilter & " and upper(借阅人) =upper([5])"
    End If
    
    If Trim(strBorrowId) <> "" Then
        strFilter = " id=[3]"
    End If
    
    '判断归档数据是否转移
    mblnMoved = MovedByDate(dtStartDate)
    
    strSql = "select /*+ Rule*/ id,id as 借阅号,借阅人,TO_CHAR(借阅时间,'yyyy-mm-dd hh:mm:ss') as 借阅时间,TO_CHAR((借阅时间 + 借阅天数),'yyyy-mm-dd hh:mm:ss') as 归还日期, 证件类型,证件号码,联系电话,联系地址,押金,借阅类型,借阅天数,借阅原因,登记人,归还状态,备注,确认状态 From 病理借阅信息 " & _
            " Where " & strFilter & " order by id "
            
    Set ufgBorrow.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                    CDate(Format(dtStartDate, "yyyy-mm-dd 00:00:00")), _
                                                    CDate(Format(dtEndDate, "yyyy-mm-dd 23:59:59")), _
                                                    Val(strBorrowId), _
                                                    strCardNo, _
                                                    strBorrowName)
    Call FilterBorrowData
End Sub


Private Sub LoadBorrowReturnHistory(ByVal lngBorrowId As Long)
'载入借阅归还历史
    Dim strSql As String
    
    If lngBorrowId <= 0 Then Exit Sub
    

    strSql = " select id, 归还人,归还日期,退还押金,外诊医院,外诊医师,外诊意见,登记人,备注 from 病理归还信息 where 借阅ID=[1] order by 归还日期"
      
    
    Set ufgBackHistory.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBorrowId)
    Call ufgBackHistory.RefreshData
End Sub


Private Sub LoadBorrowDetail(ByVal lngBorrowId As Long)
'载入借阅明细
    Dim strSql As String
    
    If lngBorrowId <= 0 Then Exit Sub
    

    strSql = " select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '蜡块' as 材料类别," & _
            " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量, nvl(a.归还数量, 0) as 归还数量,decode(f.确认状态,0,4,1,a.归还状态) as 归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a,病理借阅信息 f " & _
            " Where c.病理医嘱id = d.病理医嘱id And b.材块id = c.材块id and e.Id=b.档案ID And a.归档id = b.ID And b.资料来源 = 1 and a.借阅id = f.id And  a.借阅id = [1] " & _
        " Union All " & _
            " select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '切片' as 材料类别, " & _
            " decode(o.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量, nvl(a.归还数量, 0) as 归还数量,decode(f.确认状态,0,4,1,a.归还状态) as 归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理制片信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a,病理借阅信息 f " & _
            " Where c.病理医嘱id = d.病理医嘱id And o.病理医嘱id = c.病理医嘱id " & _
            " and b.制片id = o.id and c.材块id= o.材块id and e.id = b.档案ID and a.归档id=b.id and b.资料来源=2 and a.借阅id = f.id and a.借阅id=[1] " & _
        " Union All " & _
            " select a.归档id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, " & _
            " decode(o.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
            " decode(o.特检细目,0,decode(o.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || q.抗体名称 || decode(o.制作类型,-1,'-补',0,'','-重' || o.制作类型) || ')' as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量, nvl(a.归还数量, 0) as 归还数量,decode(f.确认状态,0,4,1,a.归还状态) as 归还状态, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理抗体信息 q, 病理特检信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a,病理借阅信息 f" & _
            " Where c.病理医嘱id = d.病理医嘱id And q.抗体ID = o.抗体ID And o.病理医嘱id = c.病理医嘱id " & _
            " and b.特检id = o.id and e.id = b.档案ID and a.归档id=b.id and b.资料来源=3 and a.借阅id = f.id and a.借阅id=[1] "
      
    
'    '查询已转存的数据
'    If mblnMoved Then
'        strSql = strSql & " union all " & GetMovedDataSql(strSql)
'    End If
    
    Set ufgMaterialDetail.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBorrowId)
    Call ufgMaterialDetail.RefreshData
End Sub


Private Sub FilterBorrowData()
'过滤借阅数据
    Dim strFilter As String
    
    strFilter = ""
    If tabFilter.Selected.Index = 0 Then    '未归还，部分归还
        strFilter = "归还状态=0 or 归还状态=2"
    End If
    
    If tabFilter.Selected.Index = 1 Then    '已遗失
        strFilter = "归还状态=3"
    End If
    
    If tabFilter.Selected.Index = 2 Then    '已归还
        strFilter = "归还状态=1"
    End If
    
    ufgBorrow.AdoData.Filter = strFilter
    
    Call ufgBorrow.RefreshData
End Sub


Private Sub InitBorrowList()
'初始化借阅列表
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("借阅列表配置", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgBorrow.IsCopyMode = True
    ufgBorrow.IsKeepRows = False
    ufgBorrow.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowCols)
        '设置行数
    ufgBorrow.GridRows = glngStandardRowCount
    '设置行高
    ufgBorrow.RowHeightMin = glngStandardRowHeight
    ufgBorrow.DefaultColNames = gstrMaterialBorrowCols
    ufgBorrow.ColConvertFormat = gstrMaterialBorrowConvertFormat
End Sub


Private Sub InitBorrowDetailList()
'初始化借阅明细列表
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("借阅明细列表配置", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgMaterialDetail.IsKeepRows = False
    ufgMaterialDetail.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowDetailCols)
        '设置行数
   ' ufgMaterialDetail.GridRows = glngStandardRowCount
    '设置行高
    ufgMaterialDetail.RowHeightMin = glngStandardRowHeight
    ufgMaterialDetail.DefaultColNames = gstrMaterialBorrowDetailCols
    ufgMaterialDetail.ColConvertFormat = gstrMaterialBorrowDetailConvertFormat
End Sub


Private Sub InitBackHistoryList()
'初始化归还历史列表
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("借阅归还列表配置", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgBackHistory.IsKeepRows = False
    ufgBackHistory.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowBackCols)
        '设置行数
    'ufgBackHistory.GridRows = glngStandardRowCount
    
    '禁止右键弹出列表配置窗口
    ufgBackHistory.IsEjectConfig = False
    '设置行高
    ufgBackHistory.RowHeightMin = glngStandardRowHeight
    ufgBackHistory.DefaultColNames = gstrMaterialBorrowBackCols
    ufgBackHistory.ColConvertFormat = gstrMaterialBorrowBackConvertFormat
End Sub


Private Sub InitTabs()
    With tabFilter
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        

        .InsertItem 0, "未归还", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "未归还"
        
        .InsertItem 1, "已遗失", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "已遗失"
        
        .InsertItem 2, "已归还", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "已归还"
        
        .InsertItem 3, "所 有", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "所 有"
        
        .Item(0).Selected = True
    End With
    
End Sub


Private Sub Form_Resize()
On Error Resume Next
    AdjustLayOut
err.Clear
End Sub


Private Sub mnu_About_Click()
'关于
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_BBS_Click()
'中联论坛
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_CancelBorrow_Click()
'撤销借阅
On Error GoTo errHandle
    Call Execute_CancelBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_DelBorrow_Click()
'删除借阅
On Error GoTo errHandle
    Call Execute_DelBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Exit_Click()
'退出
On Error GoTo errHandle
    Call Execute_Exit
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ExportExcel_Click()
'导处Excel
On Error GoTo errHandle
    Call MenuPrint(3)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Font_Click()
'字体
On Error GoTo errHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgBorrow.DataGrid.Font.Bold
    diaFont.FontName = ufgBorrow.DataGrid.Font.Name
    diaFont.FontSize = ufgBorrow.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgBorrow.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgBorrow.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '借阅列表
    ufgBorrow.DataGrid.Font.Bold = diaFont.FontBold
    ufgBorrow.DataGrid.Font.Name = diaFont.FontName
    ufgBorrow.DataGrid.Font.Size = diaFont.FontSize
    ufgBorrow.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgBorrow.DataGrid.Font.Underline = diaFont.FontUnderline
    
    
    Call ufgBorrow.DataGrid.Refresh
    
    ufgBorrow.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgBorrow.DataGrid.AutoSize(0, ufgBorrow.DataGrid.Rows - 1)
    
    ufgBorrow.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgBorrow.DataGrid.AutoSize(0, ufgBorrow.DataGrid.Rows - 1)
    
    
    '借阅材料明细列表
    ufgMaterialDetail.DataGrid.Font.Bold = diaFont.FontBold
    ufgMaterialDetail.DataGrid.Font.Name = diaFont.FontName
    ufgMaterialDetail.DataGrid.Font.Size = diaFont.FontSize
    ufgMaterialDetail.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgMaterialDetail.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgMaterialDetail.DataGrid.Refresh
    
    ufgMaterialDetail.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgMaterialDetail.DataGrid.AutoSize(0, ufgMaterialDetail.DataGrid.Rows - 1)
    
    ufgMaterialDetail.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgMaterialDetail.DataGrid.AutoSize(0, ufgMaterialDetail.DataGrid.Rows - 1)
    
    
    '归还历史列表
    ufgBackHistory.DataGrid.Font.Bold = diaFont.FontBold
    ufgBackHistory.DataGrid.Font.Name = diaFont.FontName
    ufgBackHistory.DataGrid.Font.Size = diaFont.FontSize
    ufgBackHistory.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgBackHistory.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgBackHistory.DataGrid.Refresh
    
    ufgBackHistory.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgBackHistory.DataGrid.AutoSize(0, ufgBackHistory.DataGrid.Rows - 1)
    
    ufgBackHistory.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgBackHistory.DataGrid.AutoSize(0, ufgBackHistory.DataGrid.Rows - 1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HelpMan_Click()
'帮助
On Error GoTo errHandle
    Call Execute_Help
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HomePage_Click()
'中联主页
On Error GoTo errHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_NewBorrow_Click()
'新增借阅
On Error GoTo errHandle
    Call Execute_NewBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ParameterConfig_Click()
'参数配置
On Error GoTo errHandle
    Call Execute_ParameterConfig
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Preview_Click()
'预览数据列表
On Error GoTo errHandle
    Call MenuPrint(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PreviewBorrow_Click()
'预览借阅回执
On Error GoTo errHandle
    Call Execute_PrintBorrowReceipt(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Print_Click()
'打印数据列表
On Error GoTo errHandle
    Call MenuPrint(1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintBorrow_Click()
'打印借阅回执
On Error GoTo errHandle
    Call Execute_PrintBorrowReceipt(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Execute_PrintBorrowReceipt(ByVal blnIsAtOncePrint As Boolean)
'预览打印借阅回执
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Trim(mstrBorrowReportName) = "" Then
        Call MsgBoxD(Me, "尚未配置回执单对应的报表名称，请到“参数设置”中进行配置。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, mstrBorrowReportName, Me, "借阅ID=" & Val(ufgBorrow.KeyValue(ufgBorrow.SelectionRow)), IIf(blnIsAtOncePrint, 2, 1)) '1：预览，2：打印
End Sub

Private Sub mnu_PrintConfig_Click()
'打印配置
On Error GoTo errHandle
    Call zlPrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_QueryBorrow_Click()
'借阅查询
On Error GoTo errHandle
    Call Execute_QueryBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ReturnBorrow_Click()
'借阅归还
On Error GoTo errHandle
    Call Execute_ReturnBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Send_Click()
'发送反馈
On Error GoTo errHandle
    Call zlMailTo(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StandardButton_Click()
On Error GoTo errHandle
    Dim intCount As Long
    Me.mnu_StandardButton.Checked = Not Me.mnu_StandardButton.Checked
    Me.tbrTools.Visible = Me.mnu_StandardButton.Checked
    
    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If

    Me.tbrTools.Refresh
    
    Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StateBar_Click()
On Error GoTo errHandle
    Me.mnu_StateBar.Checked = Not Me.mnu_StateBar.Checked
    Me.stbThis.Visible = Me.mnu_StateBar.Checked
    
    Call Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_UpdateBorrow_Click()
'更新借阅
On Error GoTo errHandle
    Call Execute_UpdateBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_WordLabel_Click()
On Error GoTo errHandle
    Dim intCount As Long
    
    Me.mnu_WordLabel.Checked = Not Me.mnu_WordLabel.Checked

    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If
    
    Me.tbrTools.Refresh
    
    Call Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
    tabFilter.Left = 0
    tabFilter.Top = 0
    tabFilter.Width = Picture1.ScaleWidth

    Picture3.Top = tabFilter.Height + 120
    Picture3.Left = 120
    Picture3.Width = Picture1.ScaleWidth - 120
    Picture3.Height = Picture1.ScaleHeight - 120
    
    Call ucSplitter2.RePaint
err.Clear
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    framMaterialDetail.Left = 0
    framMaterialDetail.Top = 0
    framMaterialDetail.Width = Picture2.ScaleWidth
    framMaterialDetail.Height = Picture2.ScaleHeight
    
    labBorrowDetail.Left = 120
    
    ufgMaterialDetail.Left = 120
    ufgMaterialDetail.Top = labBorrowDetail.Top + labBorrowDetail.Height
    ufgMaterialDetail.Width = framMaterialDetail.Width - 240
    ufgMaterialDetail.Height = framMaterialDetail.Height - ufgBackHistory.Height - labBackHistory.Height - labBorrowDetail.Height - 480

    labBackHistory.Left = 120
    labBackHistory.Top = ufgMaterialDetail.Top + ufgMaterialDetail.Height + 120
    
    ufgBackHistory.Left = 120
    ufgBackHistory.Top = labBackHistory.Top + labBackHistory.Height
    ufgBackHistory.Width = framMaterialDetail.Width - 240
err.Clear
End Sub

Private Sub tbn_SureBorrow_Click()
'确认借阅
On Error GoTo errHandle
    Call Execute_SureBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'过滤借阅数据
On Error GoTo errHandle
    If Not tabFilter.Visible Then Exit Sub
    
    Call FilterBorrowData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errHandle
    Select Case UCase(Button.Key)
        Case UCase("tbn_NewLend")   '新增借阅
            Call Execute_NewBorrow
            
        Case UCase("tbn_DelLend")   '删除借阅
            Call Execute_DelBorrow
            
        Case UCase("tbn_UpdateLend")    '更新借阅
            Call Execute_UpdateBorrow
        
        Case UCase("tbn_SureLend")  '确认借阅
            Call Execute_SureBorrow
            
        Case UCase("tbn_CancelLend")    '撤销借阅
            Call Execute_CancelBorrow
            
        Case UCase("tbn_QueryLend")     '查询借阅
            Call Execute_QueryBorrow
            
        Case UCase("tbn_ReturnLend")     '归还借阅
            Call Execute_ReturnBorrow
            
        Case UCase("tbn_PreviewBorrow")  '回执预览
            Call Execute_PrintBorrowReceipt(False)
        
        Case UCase("tbn_PrintBorrow")  '回执打印
            Call Execute_PrintBorrowReceipt(True)
            
        Case UCase("tbn_Help")     '帮助
            Call Execute_Help
            
        Case UCase("tbn_Exit")      '退出
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function AllowReturnBorrow(ByVal lngBorrowRow As Long) As String
'判断是否允许删除借阅
    AllowReturnBorrow = ""
    
    If mblnMoved Then
        AllowReturnBorrow = "数据已被转移，不能进行归还处理。"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngBorrowRow, gstrPatholCol_归还状态) = BorrowReturnState_Return Then
        AllowReturnBorrow = "该次借阅已归还，不能进行归还处理。"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngBorrowRow, gstrPatholCol_确认状态) = BorrowSureState_NoSure Then
        AllowReturnBorrow = "该次借阅未被确认，不能进行归还处理。"
        Exit Function
    End If
    
End Function

Private Sub Execute_ReturnBorrow()
'借阅归还
    Dim strInf As String
    Dim frmReturnBorrow As frmPatholReborrowReturn
    
    
    Set frmReturnBorrow = New frmPatholReborrowReturn
    
On Error GoTo errFree
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要归还的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = AllowReturnBorrow(ufgBorrow.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call frmReturnBorrow.ShowBorrowReturnWindow(ufgBorrow, Me)
    
    If frmReturnBorrow.blnIsOk Then
        Call ufgBorrow_OnSelChange
    End If

    Call Unload(frmReturnBorrow)
    Set frmReturnBorrow = Nothing
        
Exit Sub
errFree:
    Call Unload(frmReturnBorrow)
    Set frmReturnBorrow = Nothing
    
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigBorrowModifyState(ByVal blnIsSureBorrow As Boolean, ByVal blnIsReturn As Boolean)
'配置档案修改状态
'blnIsSureBorrow：是否确认借阅(true：已确认, false：未确认)
'blnIsReturn：是否归还
    Dim i As Long

    For i = 1 To tbrTools.Buttons.Count
        Select Case UCase(tbrTools.Buttons(i).Key)
            Case UCase("tbn_DelLend"), UCase("tbn_UpdateLend"), UCase("tbn_SureLend")
                tbrTools.Buttons(i).Enabled = Not blnIsSureBorrow
            Case UCase("tbn_CancelLend"), UCase("tbn_ReturnLend")
                tbrTools.Buttons(i).Enabled = blnIsSureBorrow And Not blnIsReturn
        End Select
    Next i
    
    mnu_DelBorrow.Enabled = Not blnIsSureBorrow
    mnu_UpdateBorrow.Enabled = Not blnIsSureBorrow
    mnu_DelBorrow.Enabled = Not blnIsSureBorrow
    mnu_SureBorrow.Enabled = Not blnIsSureBorrow
    
    mnu_CancelBorrow.Enabled = blnIsSureBorrow
    mnu_ReturnBorrow.Enabled = blnIsSureBorrow And Not blnIsReturn
    
    
End Sub



Private Sub Execute_NewBorrow()
'新增借阅
    Dim frmNewBorrow As frmPatholReborrowNew
    
    Set frmNewBorrow = New frmPatholReborrowNew
    On Error GoTo errFree
        Call frmNewBorrow.ShowNewBorrowWindow(ufgBorrow, Me)
        
        '读取档案附加显示信息
        If ufgBorrow.IsSelectionRow Then
            Call ReadBorrowInf(ufgBorrow.SelectionRow)
        End If
        
        Call RefreshStateInf

        Call Unload(frmNewBorrow)
        Set frmNewBorrow = Nothing
        
        Exit Sub
errFree:
    Call Unload(frmNewBorrow)
    Set frmNewBorrow = Nothing
    
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Execute_UpdateBorrow()
'更新借阅
    Dim frmUpdateBorrow As frmPatholReborrowNew
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "没有选择需要更新的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmUpdateBorrow = New frmPatholReborrowNew
    On Error GoTo errFree
        Call frmUpdateBorrow.ShowUpdateBorrowWindow(ufgBorrow, Me)
        
        '更新借阅显示
        If frmUpdateBorrow.blnIsOk Then
            Call ufgBorrow_OnSelChange
        End If

        Call Unload(frmUpdateBorrow)
        Set frmUpdateBorrow = Nothing
        
        Exit Sub
errFree:
    Call Unload(frmUpdateBorrow)
    Set frmUpdateBorrow = Nothing
    
    If ErrCenter() = 1 Then Resume
End Sub


Private Function AllowDelBorrow(ByVal lngDelRow As Long) As String
'判断是否允许删除借阅
    AllowDelBorrow = ""
    
    If mblnMoved Then
        AllowDelBorrow = "数据已被转移，不能进行删除。"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngDelRow, gstrPatholCol_确认状态) <> BorrowSureState_NoSure Then
        AllowDelBorrow = "已确认借阅，不能进行删除。"
        Exit Function
    End If
    
End Function


Private Sub Execute_DelBorrow()
'删除借阅(只有未被确认的借阅才能被删除)
    Dim strInf As String
    
    '需要判断档案是否已经封存，且档案中不包含检查
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = AllowDelBorrow(ufgBorrow.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除选择的借阅记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    Call DelArchivesBorrow(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    Call ufgBorrow.DelRow(ufgBorrow.SelectionRow, False, True)
    
    '读取档案附加显示信息
    If ufgBorrow.IsSelectionRow Then
        Call ReadBorrowInf(ufgBorrow.SelectionRow)
    Else
        '如果没有借阅记录，则情况借阅材料明细
        Call ufgMaterialDetail.ClearListData
    End If
    
    Call RefreshStateInf
End Sub


Private Sub Execute_SureBorrow()
'确认借阅
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要确认的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_确认状态) = BorrowSureState_Sure Then
        Call MsgBoxD(Me, "该次借阅已被确认，不能进行确认处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    '更新确认状态
    Call zlDatabase.ExecuteProcedure("Zl_病理借阅_确认状态更新(" & ufgBorrow.KeyValue(ufgBorrow.SelectionRow) & ",1)", Me.Caption)
    
    
    Call ufgBorrow.SyncText(ufgBorrow.SelectionRow, gstrPatholCol_确认状态, BorrowSureState_Sure, True)
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(True, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_归还状态) = "已归还")
    
    Call ufgMaterialDetail.SyncText(ufgMaterialDetail.SelectionRow, gstrPatholCol_归还状态, "未归还", False)
    
'    Call RefreshStateInf(True, False)

    '自动打印借阅回执单
    If mblnIsAutoPrint Then
        Call Execute_PrintBorrowReceipt(True)
    End If
End Sub


Private Sub Execute_QueryBorrow()
'查询借阅
    Dim strSql As String
    
    Call frmPatholReborrowQuery.ShowBorrowQueryWindow(mlngDefaultQueryDays, Me)
    
    If frmPatholReborrowQuery.mblnIsOk Then
        Call QueryBorrowData(frmPatholReborrowQuery.dtStartDate, frmPatholReborrowQuery.dtEndDate, _
            frmPatholReborrowQuery.strBorrowId, frmPatholReborrowQuery.strCardNo, frmPatholReborrowQuery.strBorrowName)
    End If
End Sub

Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕打印预览, 0弹出操作选择对话框，1预览，2打印，3导出Excel
    '参数：    输出方式
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgBorrow.DataGrid
    
    objPrint.Title = "病理借阅清单"

    If intOutMode = 0 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrView1Grd objPrint, intOutMode
    End If

End Sub


Private Sub Execute_Exit()
'退出
    Call Unload(Me)
End Sub

Private Sub Execute_Help()
'帮助
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Execute_ParameterConfig()
'参数配置
    Dim frmParameter As frmPatholReborrowParameter
    
    Set frmParameter = New frmPatholReborrowParameter
On Error GoTo errFree
    Call frmParameter.ShowParameterWindow(mlngDefaultQueryDays, mstrBorrowReportName, mblnIsAutoPrint, Me)
    
    mlngDefaultQueryDays = frmParameter.lngDefaultQueryDays
    mstrBorrowReportName = frmParameter.strLabelReportName
    mblnIsAutoPrint = frmParameter.blnIsAutoPrint
    
errFree:
    Call Unload(frmParameter)
    Set frmParameter = Nothing
    
End Sub


Private Sub Execute_CancelBorrow()
'撤销借阅
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要撤销的借阅记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_确认状态) = BorrowSureState_NoSure Then
        Call MsgBoxD(Me, "该次借阅未被确认，不能进行撤销处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBackHistory.ShowingDataRowCount > 0 Then
        Call MsgBoxD(Me, "该次借阅存在归还历史记录，不能进行撤销处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要对该借阅进行撤销操作吗？借阅撤销后相关信息将允许被修改。", vbYesNo, Me.Caption) = vbNo Then Exit Sub

    '更新确认状态
    Call zlDatabase.ExecuteProcedure("Zl_病理借阅_确认状态更新(" & ufgBorrow.KeyValue(ufgBorrow.SelectionRow) & ",0)", Me.Caption)
    
    
    Call ufgBorrow.SyncText(ufgBorrow.SelectionRow, gstrPatholCol_确认状态, BorrowSureState_NoSure, True)
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(False, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_归还状态) = "已归还")
    
    Call ufgMaterialDetail.SyncText(ufgMaterialDetail.SelectionRow, gstrPatholCol_归还状态, "待借出", False)
    
'    Call RefreshStateInf(True, False)
End Sub

Private Sub DelArchivesBorrow(ByVal lngBorrowId As Long)
'删除借阅
    Dim strSql As String
    
    strSql = "Zl_病理借阅_删除借阅(" & lngBorrowId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub


Private Sub ufgBackHistory_OnColFormartChange()
'保存借阅归还列表的列配置
On Error GoTo errHandle
    zlDatabase.SetPara "借阅归还列表配置", ufgBackHistory.GetColsString(ufgBackHistory), glngSys, G_LNG_PATHOLBORROW_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgBorrow_OnColFormartChange()
'保存借阅列表的列配置
On Error GoTo errHandle
    zlDatabase.SetPara "借阅列表配置", ufgBorrow.GetColsString(ufgBorrow), glngSys, G_LNG_PATHOLBORROW_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgBorrow_OnColsNameReSet()
On Error GoTo errHandle
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Call QueryBorrowData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgBorrow_OnNewRow(ByVal Row As Long)
On Error GoTo errHandle
    If CDate(ufgBorrow.Text(Row, gstrPatholCol_归还日期)) < mdtServicesTime And ufgBorrow.Text(Row, gstrPatholCol_归还状态) <> "已归还" Then
        ufgBorrow.DataGrid.Cell(flexcpPicture, Row, ufgBorrow.GetColIndex(gstrPatholCol_借阅号)) = picTimeOut
    End If
Exit Sub
errHandle:
If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgBorrow_OnSelChange()
On Error GoTo errHandle
    If Not ufgBorrow.IsSelectionRow Then
        Exit Sub
    End If
    
    If ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_ID) = "" Then Exit Sub
    
    '载入借阅明细
    Call LoadBorrowDetail(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    '载入归还历史
    Call LoadBorrowReturnHistory(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_确认状态) = BorrowSureState_Sure, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_归还状态) = "已归还")
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMaterialDetail_OnColFormartChange()
'保存借阅明细列表的列配置
On Error GoTo errHandle
    zlDatabase.SetPara "借阅明细列表配置", ufgMaterialDetail.GetColsString(ufgMaterialDetail), glngSys, G_LNG_PATHOLBORROW_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgMaterialDetail_OnColsNameReSet()
On Error GoTo errHandle

    '载入借阅明细
    Call LoadBorrowDetail(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
