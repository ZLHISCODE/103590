VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholReborrowStation 
   Caption         =   "����軹����վ"
   ClientHeight    =   9555
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13500
   Icon            =   "frmPatholReborrowStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   13500
   StartUpPosition =   3  '����ȱʡ
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
               Name            =   "����"
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
            Caption         =   "�黹��ʷ��"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   5160
            Width           =   975
         End
         Begin VB.Label labBorrowDetail 
            Caption         =   "���Ĳ��ϣ�"
            BeginProperty Font 
               Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "δ�黹������"
            TextSave        =   "δ�黹������"
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "����δ�黹������"
            TextSave        =   "����δ�黹������"
            Key             =   "sb_NoEnterWaxStone"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
            Text            =   "���ֹ黹������"
            TextSave        =   "���ֹ黹������"
            Key             =   "sb_NoEnterSlices"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "����ʧ����������"
            TextSave        =   "����ʧ����������"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
            Caption         =   "��ִԤ��"
            Key             =   "tbn_PreviewBorrow"
            Object.Tag             =   "��ִԤ��"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ִ��ӡ"
            Key             =   "tbn_PrintBorrow"
            Object.Tag             =   "��ִ��ӡ"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "tbn_NewLend"
            Object.Tag             =   "��������"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ������"
            Key             =   "tbn_DelLend"
            Object.Tag             =   "ɾ������"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���½���"
            Key             =   "tbn_UpdateLend"
            Object.Tag             =   "���½���"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѯ����"
            Key             =   "tbn_QueryLend"
            Object.Tag             =   "��ѯ����"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȷ�Ͻ���"
            Key             =   "tbn_SureLend"
            Object.Tag             =   "ȷ�Ͻ���"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "tbn_CancelLend"
            Object.Tag             =   "��������"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���Ĺ黹"
            Key             =   "tbn_ReturnLend"
            Object.Tag             =   "���Ĺ黹"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "tbn_Help"
            Object.Tag             =   "����"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "tbn_Exit"
            Object.Tag             =   "�˳�"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnu_Flie 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnu_ParameterConfig 
         Caption         =   "��������(&C)"
      End
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "��ӡ����(&N)"
      End
      Begin VB.Menu mnu_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Preview 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "��ӡ(&P)"
      End
      Begin VB.Menu mnu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportExcel 
         Caption         =   "�����Execl(&E)"
      End
      Begin VB.Menu mnu_Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "�˳�(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnu_NewBorrow 
         Caption         =   "��������(&N)"
      End
      Begin VB.Menu mnu_DelBorrow 
         Caption         =   "ɾ������(&D)"
      End
      Begin VB.Menu mnu_UpdateBorrow 
         Caption         =   "���½���(&U)"
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SureBorrow 
         Caption         =   "ȷ�Ͻ���(&S)"
      End
      Begin VB.Menu mnu_CancelBorrow 
         Caption         =   "��������(&C)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_QueryBorrow 
         Caption         =   "���Ĳ�ѯ(&Q)"
      End
      Begin VB.Menu mnu_Split6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ReturnBorrow 
         Caption         =   "���Ĺ黹(&R)"
      End
      Begin VB.Menu mnu_Split9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PreviewBorrow 
         Caption         =   "��ִԤ��(&V)"
      End
      Begin VB.Menu mnu_PrintBorrow 
         Caption         =   "��ִ��ӡ(&P)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnu_ToolBar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnu_StandardButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_WordLabel 
            Caption         =   "�ı���ǩ(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_StateBar 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Split7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Font 
         Caption         =   "����(&F)"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "����(&T)"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "����(&H)"
      Begin VB.Menu mnu_HelpMan 
         Caption         =   "��������(&H)"
      End
      Begin VB.Menu mnu_Web 
         Caption         =   "WEB�ϵ�����(&W)"
         Begin VB.Menu mnu_HomePage 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnu_bbs 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnu_Send 
            Caption         =   "���ͷ���(&K)"
         End
      End
      Begin VB.Menu mnu_Split8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "����...(&A)"
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

'ȷ��״̬
Private Const BorrowSureState_NoSure As String = "δȷ��"
Private Const BorrowSureState_Sure As String = "��ȷ��"

'�黹״̬
Private Const BorrowReturnState_Return As String = "�ѹ黹" 'δ�黹�����ֹ黹����ʧ����

'Ϊ�˵�������Ӧ��ͼ��
Private Const MF_BITMAP = &H400&


Private mstrPrivs As String
Private mlngDefaultQueryDays As Long
Private mstrBorrowReportName As String
Private mblnIsAutoPrint As Boolean
Private mblnMoved As Boolean

Private mdtServicesTime As Date

Private WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long 'ȡ�ô��ڵĲ˵����,hwnd�Ǵ��ڵľ��
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal npos As Long) As Long 'ȡ���Ӳ˵������nPos�ǲ˵���λ��
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal npos As Long, ByVal wFlags As Long, ByVal hBitUnchecked As Long, ByVal hBitChecked As Long) As Long




Private Sub AdjustLayOut()
    
    Picture1.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    Picture1.Height = Me.ScaleHeight - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
    
    Call ucSplitter1.RePaint
End Sub


Private Sub LoadParameterConfig()
'������ز�������
    mlngDefaultQueryDays = zlDatabase.GetPara("����Ĭ�ϲ�ѯ����", glngSys, G_LNG_PATHOLBORROW_NUM, "100")
    mstrBorrowReportName = zlDatabase.GetPara("���Ļ�ִ��������", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    mblnIsAutoPrint = zlDatabase.GetPara("����ȷ�Ϻ��Զ���ӡ��ִ", glngSys, G_LNG_PATHOLBORROW_NUM, 1)
End Sub



Private Sub ConfigPopedomFace()
'����Ȩ�����ý��棬������߱�Ȩ��ʱ�������ض�Ӧ���ܰ�ť
    Dim i As Long
    
    mnu_ParameterConfig.Visible = CheckPopedom(mstrPrivs, "��������")
    
    mnu_CancelBorrow.Visible = CheckPopedom(mstrPrivs, "��������")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_CancelLend") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "��������")
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
'ˢ�²�����ʧ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select " & _
            " sum(case when �黹״̬=0 then 1 else 0 end)  as δ�黹����, " & _
            " sum(case when (����ʱ��+��������<sysdate and �黹״̬=0) then 1 else 0 end)  as ����δ�黹����, " & _
            " sum(case when �黹״̬=2 then 1 else 0 end)as ���ֹ黹����, " & _
            " sum(case when �黹״̬=3 then 1 else 0 end) as ��ʧ���� " & _
            " From ���������Ϣ "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If rsData.RecordCount > 0 Then
        stbThis.Panels(2).Text = "δ�黹������" & Nvl(rsData!δ�黹����)
        stbThis.Panels(3).Text = "����δ�黹������" & Nvl(rsData!����δ�黹����)
        stbThis.Panels(4).Text = "���ֹ黹������" & Nvl(rsData!���ֹ黹����)
        stbThis.Panels(5).Text = "����ʧ����������" & Nvl(rsData!��ʧ����)
    End If
End Sub


Private Sub InitMenuIcoConfig()
'��ʼ���˵�ͼ����ʾ
On Error Resume Next
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim hSubSubMenu As Long
    
    hMenu = GetMenu(Me.hWnd)
    
    '���õ�һ��˵�(�ļ�)
    hSubMenu = GetSubMenu(hMenu, 0) 'ȡ�õ�һ��˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(28).Picture, imgMenus.ListImages(28).Picture) '��������
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(3).Picture, imgMenus.ListImages(3).Picture) '��ӡ����
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '��ӡԤ��
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '��ӡ
    Call SetMenuItemBitmaps(hSubMenu, 6, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '����Excel
    Call SetMenuItemBitmaps(hSubMenu, 8, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '�˳�
    

    '���õڶ���˵����༭��
    hSubMenu = GetSubMenu(hMenu, 1) 'ȡ�õڶ���˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '��������
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(8).Picture, imgMenus.ListImages(8).Picture) 'ɾ������
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(9).Picture, imgMenus.ListImages(9).Picture) '���½���
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(11).Picture, imgMenus.ListImages(11).Picture) 'ȷ�Ͻ���
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '��������
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(12).Picture, imgMenus.ListImages(12).Picture) '��ѯ����
    Call SetMenuItemBitmaps(hSubMenu, 9, MF_BITMAP, imgMenus.ListImages(29).Picture, imgMenus.ListImages(29).Picture) '���Ĺ黹
    Call SetMenuItemBitmaps(hSubMenu, 11, MF_BITMAP, imgMenus.ListImages(1).Picture, imgMenus.ListImages(1).Picture) '��ִԤ��
    Call SetMenuItemBitmaps(hSubMenu, 12, MF_BITMAP, imgMenus.ListImages(2).Picture, imgMenus.ListImages(2).Picture) '��ִ��ӡ
    
    
    '���õڶ���˵����鿴��
    hSubMenu = GetSubMenu(hMenu, 2) 'ȡ�õڶ���˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(27).Picture, imgMenus.ListImages(27).Picture) '������
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(22).Picture, imgMenus.ListImages(21).Picture) '״̬��
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(23).Picture, imgMenus.ListImages(23).Picture) '����
    
        hSubSubMenu = GetSubMenu(hSubMenu, 0)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(26).Picture, imgMenus.ListImages(20).Picture) '��׼��ť
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(25).Picture, imgMenus.ListImages(24).Picture) '�ı���ǩ
    
    
    
    '���õ�����˵���������
    hSubMenu = GetSubMenu(hMenu, 3) 'ȡ�õ�����˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(13).Picture, imgMenus.ListImages(13).Picture) '��������
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(14).Picture) 'web����
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(15).Picture, imgMenus.ListImages(15).Picture) '��
    
        hSubSubMenu = GetSubMenu(hSubMenu, 1)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(13).Picture) '��������
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(16).Picture, imgMenus.ListImages(16).Picture) '������̳
        Call SetMenuItemBitmaps(hSubSubMenu, 2, MF_BITMAP, imgMenus.ListImages(17).Picture, imgMenus.ListImages(17).Picture) '���ͷ���
    
    err.Clear

End Sub



Private Sub ReadBorrowInf(ByVal lngArchivesRowIndex As Long)
'��ȡ������Ϣ
    Dim strInf As String
    If lngArchivesRowIndex <= 0 Then Exit Sub
    
    strInf = "�� �� �ţ�" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_���ĺ�) & vbCrLf
    strInf = strInf & "�� �� �ˣ�" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_������) & vbCrLf
    strInf = strInf & "֤�����ͣ�" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_֤������) & vbCrLf
    strInf = strInf & "֤�����룺" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_֤������) & vbCrLf
    strInf = strInf & "��ϵ�绰��" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_��ϵ�绰) & vbCrLf
    strInf = strInf & "��ϵ��ַ��" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_��ϵ��ַ) & vbCrLf
    strInf = strInf & "����Ѻ��" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_Ѻ��) & vbCrLf
    strInf = strInf & "�������ڣ�" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "����������" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "�黹���ڣ�" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_�黹����) & vbCrLf
    strInf = strInf & "����ԭ��" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_����ԭ��) & vbCrLf
    strInf = strInf & "�黹״̬��" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_�黹״̬) & vbCrLf
    strInf = strInf & "��ע˵����" & ufgBorrow.Text(lngArchivesRowIndex, gstrPatholCol_��ע)
    
    rtbDetail.Text = strInf
End Sub



Private Sub QueryBorrowData(ByVal dtStartDate As Date, ByVal dtEndDate As Date, _
    Optional ByVal strBorrowId As String = "", Optional ByVal strCardNo As String = "", _
    Optional ByVal strBorrowName As String = "")
'��ѯ���ļ�¼����
    Dim strSql As String
    Dim strFilter As String
    
    
    mdtServicesTime = zlDatabase.Currentdate
    
    strFilter = ""
    
    strFilter = " ����ʱ�� between [1] and [2]"
    
    If Trim(strCardNo) <> "" Then
        strFilter = strFilter & " and upper(֤������)=upper([4])"
    End If
    
    If Trim(strBorrowName) <> "" Then
        strFilter = strFilter & " and upper(������) =upper([5])"
    End If
    
    If Trim(strBorrowId) <> "" Then
        strFilter = " id=[3]"
    End If
    
    '�жϹ鵵�����Ƿ�ת��
    mblnMoved = MovedByDate(dtStartDate)
    
    strSql = "select /*+ Rule*/ id,id as ���ĺ�,������,TO_CHAR(����ʱ��,'yyyy-mm-dd hh:mm:ss') as ����ʱ��,TO_CHAR((����ʱ�� + ��������),'yyyy-mm-dd hh:mm:ss') as �黹����, ֤������,֤������,��ϵ�绰,��ϵ��ַ,Ѻ��,��������,��������,����ԭ��,�Ǽ���,�黹״̬,��ע,ȷ��״̬ From ���������Ϣ " & _
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
'������Ĺ黹��ʷ
    Dim strSql As String
    
    If lngBorrowId <= 0 Then Exit Sub
    

    strSql = " select id, �黹��,�黹����,�˻�Ѻ��,����ҽԺ,����ҽʦ,�������,�Ǽ���,��ע from ����黹��Ϣ where ����ID=[1] order by �黹����"
      
    
    Set ufgBackHistory.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBorrowId)
    Call ufgBackHistory.RefreshData
End Sub


Private Sub LoadBorrowDetail(ByVal lngBorrowId As Long)
'���������ϸ
    Dim strSql As String
    
    If lngBorrowId <= 0 Then Exit Sub
    

    strSql = " select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '����' as �������," & _
            " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������, nvl(a.�黹����, 0) as �黹����,decode(f.ȷ��״̬,0,4,1,a.�黹״̬) as �黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a,���������Ϣ f " & _
            " Where c.����ҽ��id = d.����ҽ��id And b.�Ŀ�id = c.�Ŀ�id and e.Id=b.����ID And a.�鵵id = b.ID And b.������Դ = 1 and a.����id = f.id And  a.����id = [1] " & _
        " Union All " & _
            " select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '��Ƭ' as �������, " & _
            " decode(o.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������, nvl(a.�黹����, 0) as �黹����,decode(f.ȷ��״̬,0,4,1,a.�黹״̬) as �黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ������Ƭ��Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a,���������Ϣ f " & _
            " Where c.����ҽ��id = d.����ҽ��id And o.����ҽ��id = c.����ҽ��id " & _
            " and b.��Ƭid = o.id and c.�Ŀ�id= o.�Ŀ�id and e.id = b.����ID and a.�鵵id=b.id and b.������Դ=2 and a.����id = f.id and a.����id=[1] " & _
        " Union All " & _
            " select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, " & _
            " decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
            " decode(o.�ؼ�ϸĿ,0,decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || q.�������� || decode(o.��������,-1,'-��',0,'','-��' || o.��������) || ')' as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������, nvl(a.�黹����, 0) as �黹����,decode(f.ȷ��״̬,0,4,1,a.�黹״̬) as �黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ q, �����ؼ���Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a,���������Ϣ f" & _
            " Where c.����ҽ��id = d.����ҽ��id And q.����ID = o.����ID And o.����ҽ��id = c.����ҽ��id " & _
            " and b.�ؼ�id = o.id and e.id = b.����ID and a.�鵵id=b.id and b.������Դ=3 and a.����id = f.id and a.����id=[1] "
      
    
'    '��ѯ��ת�������
'    If mblnMoved Then
'        strSql = strSql & " union all " & GetMovedDataSql(strSql)
'    End If
    
    Set ufgMaterialDetail.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBorrowId)
    Call ufgMaterialDetail.RefreshData
End Sub


Private Sub FilterBorrowData()
'���˽�������
    Dim strFilter As String
    
    strFilter = ""
    If tabFilter.Selected.Index = 0 Then    'δ�黹�����ֹ黹
        strFilter = "�黹״̬=0 or �黹״̬=2"
    End If
    
    If tabFilter.Selected.Index = 1 Then    '����ʧ
        strFilter = "�黹״̬=3"
    End If
    
    If tabFilter.Selected.Index = 2 Then    '�ѹ黹
        strFilter = "�黹״̬=1"
    End If
    
    ufgBorrow.AdoData.Filter = strFilter
    
    Call ufgBorrow.RefreshData
End Sub


Private Sub InitBorrowList()
'��ʼ�������б�
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("�����б�����", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgBorrow.IsCopyMode = True
    ufgBorrow.IsKeepRows = False
    ufgBorrow.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowCols)
        '��������
    ufgBorrow.GridRows = glngStandardRowCount
    '�����и�
    ufgBorrow.RowHeightMin = glngStandardRowHeight
    ufgBorrow.DefaultColNames = gstrMaterialBorrowCols
    ufgBorrow.ColConvertFormat = gstrMaterialBorrowConvertFormat
End Sub


Private Sub InitBorrowDetailList()
'��ʼ��������ϸ�б�
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("������ϸ�б�����", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgMaterialDetail.IsKeepRows = False
    ufgMaterialDetail.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowDetailCols)
        '��������
   ' ufgMaterialDetail.GridRows = glngStandardRowCount
    '�����и�
    ufgMaterialDetail.RowHeightMin = glngStandardRowHeight
    ufgMaterialDetail.DefaultColNames = gstrMaterialBorrowDetailCols
    ufgMaterialDetail.ColConvertFormat = gstrMaterialBorrowDetailConvertFormat
End Sub


Private Sub InitBackHistoryList()
'��ʼ���黹��ʷ�б�
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara("���Ĺ黹�б�����", glngSys, G_LNG_PATHOLBORROW_NUM, "")
    
    ufgBackHistory.IsKeepRows = False
    ufgBackHistory.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialBorrowBackCols)
        '��������
    'ufgBackHistory.GridRows = glngStandardRowCount
    
    '��ֹ�Ҽ������б����ô���
    ufgBackHistory.IsEjectConfig = False
    '�����и�
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
        

        .InsertItem 0, "δ�黹", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "δ�黹"
        
        .InsertItem 1, "����ʧ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "����ʧ"
        
        .InsertItem 2, "�ѹ黹", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�ѹ黹"
        
        .InsertItem 3, "�� ��", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�� ��"
        
        .Item(0).Selected = True
    End With
    
End Sub


Private Sub Form_Resize()
On Error Resume Next
    AdjustLayOut
err.Clear
End Sub


Private Sub mnu_About_Click()
'����
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_BBS_Click()
'������̳
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_CancelBorrow_Click()
'��������
On Error GoTo errHandle
    Call Execute_CancelBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_DelBorrow_Click()
'ɾ������
On Error GoTo errHandle
    Call Execute_DelBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Exit_Click()
'�˳�
On Error GoTo errHandle
    Call Execute_Exit
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ExportExcel_Click()
'����Excel
On Error GoTo errHandle
    Call MenuPrint(3)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Font_Click()
'����
On Error GoTo errHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgBorrow.DataGrid.Font.Bold
    diaFont.FontName = ufgBorrow.DataGrid.Font.Name
    diaFont.FontSize = ufgBorrow.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgBorrow.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgBorrow.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '�����б�
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
    
    
    '���Ĳ�����ϸ�б�
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
    
    
    '�黹��ʷ�б�
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
'����
On Error GoTo errHandle
    Call Execute_Help
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HomePage_Click()
'������ҳ
On Error GoTo errHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_NewBorrow_Click()
'��������
On Error GoTo errHandle
    Call Execute_NewBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ParameterConfig_Click()
'��������
On Error GoTo errHandle
    Call Execute_ParameterConfig
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Preview_Click()
'Ԥ�������б�
On Error GoTo errHandle
    Call MenuPrint(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PreviewBorrow_Click()
'Ԥ�����Ļ�ִ
On Error GoTo errHandle
    Call Execute_PrintBorrowReceipt(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Print_Click()
'��ӡ�����б�
On Error GoTo errHandle
    Call MenuPrint(1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintBorrow_Click()
'��ӡ���Ļ�ִ
On Error GoTo errHandle
    Call Execute_PrintBorrowReceipt(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Execute_PrintBorrowReceipt(ByVal blnIsAtOncePrint As Boolean)
'Ԥ����ӡ���Ļ�ִ
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�Ľ��ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Trim(mstrBorrowReportName) = "" Then
        Call MsgBoxD(Me, "��δ���û�ִ����Ӧ�ı������ƣ��뵽���������á��н������á�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call zlReport.ReportOpen(gcnOracle, 100, mstrBorrowReportName, Me, "����ID=" & Val(ufgBorrow.KeyValue(ufgBorrow.SelectionRow)), IIf(blnIsAtOncePrint, 2, 1)) '1��Ԥ����2����ӡ
End Sub

Private Sub mnu_PrintConfig_Click()
'��ӡ����
On Error GoTo errHandle
    Call zlPrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_QueryBorrow_Click()
'���Ĳ�ѯ
On Error GoTo errHandle
    Call Execute_QueryBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ReturnBorrow_Click()
'���Ĺ黹
On Error GoTo errHandle
    Call Execute_ReturnBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Send_Click()
'���ͷ���
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
'���½���
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
'ȷ�Ͻ���
On Error GoTo errHandle
    Call Execute_SureBorrow
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���˽�������
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
        Case UCase("tbn_NewLend")   '��������
            Call Execute_NewBorrow
            
        Case UCase("tbn_DelLend")   'ɾ������
            Call Execute_DelBorrow
            
        Case UCase("tbn_UpdateLend")    '���½���
            Call Execute_UpdateBorrow
        
        Case UCase("tbn_SureLend")  'ȷ�Ͻ���
            Call Execute_SureBorrow
            
        Case UCase("tbn_CancelLend")    '��������
            Call Execute_CancelBorrow
            
        Case UCase("tbn_QueryLend")     '��ѯ����
            Call Execute_QueryBorrow
            
        Case UCase("tbn_ReturnLend")     '�黹����
            Call Execute_ReturnBorrow
            
        Case UCase("tbn_PreviewBorrow")  '��ִԤ��
            Call Execute_PrintBorrowReceipt(False)
        
        Case UCase("tbn_PrintBorrow")  '��ִ��ӡ
            Call Execute_PrintBorrowReceipt(True)
            
        Case UCase("tbn_Help")     '����
            Call Execute_Help
            
        Case UCase("tbn_Exit")      '�˳�
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function AllowReturnBorrow(ByVal lngBorrowRow As Long) As String
'�ж��Ƿ�����ɾ������
    AllowReturnBorrow = ""
    
    If mblnMoved Then
        AllowReturnBorrow = "�����ѱ�ת�ƣ����ܽ��й黹����"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngBorrowRow, gstrPatholCol_�黹״̬) = BorrowReturnState_Return Then
        AllowReturnBorrow = "�ôν����ѹ黹�����ܽ��й黹����"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngBorrowRow, gstrPatholCol_ȷ��״̬) = BorrowSureState_NoSure Then
        AllowReturnBorrow = "�ôν���δ��ȷ�ϣ����ܽ��й黹����"
        Exit Function
    End If
    
End Function

Private Sub Execute_ReturnBorrow()
'���Ĺ黹
    Dim strInf As String
    Dim frmReturnBorrow As frmPatholReborrowReturn
    
    
    Set frmReturnBorrow = New frmPatholReborrowReturn
    
On Error GoTo errFree
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�黹�Ľ��ļ�¼��", vbOKOnly, Me.Caption)
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
'���õ����޸�״̬
'blnIsSureBorrow���Ƿ�ȷ�Ͻ���(true����ȷ��, false��δȷ��)
'blnIsReturn���Ƿ�黹
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
'��������
    Dim frmNewBorrow As frmPatholReborrowNew
    
    Set frmNewBorrow = New frmPatholReborrowNew
    On Error GoTo errFree
        Call frmNewBorrow.ShowNewBorrowWindow(ufgBorrow, Me)
        
        '��ȡ����������ʾ��Ϣ
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
'���½���
    Dim frmUpdateBorrow As frmPatholReborrowNew
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "û��ѡ����Ҫ���µĽ��ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmUpdateBorrow = New frmPatholReborrowNew
    On Error GoTo errFree
        Call frmUpdateBorrow.ShowUpdateBorrowWindow(ufgBorrow, Me)
        
        '���½�����ʾ
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
'�ж��Ƿ�����ɾ������
    AllowDelBorrow = ""
    
    If mblnMoved Then
        AllowDelBorrow = "�����ѱ�ת�ƣ����ܽ���ɾ����"
        Exit Function
    End If
    
    If ufgBorrow.Text(lngDelRow, gstrPatholCol_ȷ��״̬) <> BorrowSureState_NoSure Then
        AllowDelBorrow = "��ȷ�Ͻ��ģ����ܽ���ɾ����"
        Exit Function
    End If
    
End Function


Private Sub Execute_DelBorrow()
'ɾ������(ֻ��δ��ȷ�ϵĽ��Ĳ��ܱ�ɾ��)
    Dim strInf As String
    
    '��Ҫ�жϵ����Ƿ��Ѿ���棬�ҵ����в��������
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���Ľ��ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = AllowDelBorrow(ufgBorrow.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ��ѡ��Ľ��ļ�¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    Call DelArchivesBorrow(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    Call ufgBorrow.DelRow(ufgBorrow.SelectionRow, False, True)
    
    '��ȡ����������ʾ��Ϣ
    If ufgBorrow.IsSelectionRow Then
        Call ReadBorrowInf(ufgBorrow.SelectionRow)
    Else
        '���û�н��ļ�¼����������Ĳ�����ϸ
        Call ufgMaterialDetail.ClearListData
    End If
    
    Call RefreshStateInf
End Sub


Private Sub Execute_SureBorrow()
'ȷ�Ͻ���
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫȷ�ϵĽ��ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_ȷ��״̬) = BorrowSureState_Sure Then
        Call MsgBoxD(Me, "�ôν����ѱ�ȷ�ϣ����ܽ���ȷ�ϴ���", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    '����ȷ��״̬
    Call zlDatabase.ExecuteProcedure("Zl_�������_ȷ��״̬����(" & ufgBorrow.KeyValue(ufgBorrow.SelectionRow) & ",1)", Me.Caption)
    
    
    Call ufgBorrow.SyncText(ufgBorrow.SelectionRow, gstrPatholCol_ȷ��״̬, BorrowSureState_Sure, True)
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(True, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_�黹״̬) = "�ѹ黹")
    
    Call ufgMaterialDetail.SyncText(ufgMaterialDetail.SelectionRow, gstrPatholCol_�黹״̬, "δ�黹", False)
    
'    Call RefreshStateInf(True, False)

    '�Զ���ӡ���Ļ�ִ��
    If mblnIsAutoPrint Then
        Call Execute_PrintBorrowReceipt(True)
    End If
End Sub


Private Sub Execute_QueryBorrow()
'��ѯ����
    Dim strSql As String
    
    Call frmPatholReborrowQuery.ShowBorrowQueryWindow(mlngDefaultQueryDays, Me)
    
    If frmPatholReborrowQuery.mblnIsOk Then
        Call QueryBorrowData(frmPatholReborrowQuery.dtStartDate, frmPatholReborrowQuery.dtEndDate, _
            frmPatholReborrowQuery.strBorrowId, frmPatholReborrowQuery.strCardNo, frmPatholReborrowQuery.strBorrowName)
    End If
End Sub

Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��ӡԤ��, 0��������ѡ��Ի���1Ԥ����2��ӡ��3����Excel
    '������    �����ʽ
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgBorrow.DataGrid
    
    objPrint.Title = "��������嵥"

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
'�˳�
    Call Unload(Me)
End Sub

Private Sub Execute_Help()
'����
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Execute_ParameterConfig()
'��������
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
'��������
    If Not ufgBorrow.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����Ľ��ļ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_ȷ��״̬) = BorrowSureState_NoSure Then
        Call MsgBoxD(Me, "�ôν���δ��ȷ�ϣ����ܽ��г�������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgBackHistory.ShowingDataRowCount > 0 Then
        Call MsgBoxD(Me, "�ôν��Ĵ��ڹ黹��ʷ��¼�����ܽ��г�������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫ�Ըý��Ľ��г��������𣿽��ĳ����������Ϣ�������޸ġ�", vbYesNo, Me.Caption) = vbNo Then Exit Sub

    '����ȷ��״̬
    Call zlDatabase.ExecuteProcedure("Zl_�������_ȷ��״̬����(" & ufgBorrow.KeyValue(ufgBorrow.SelectionRow) & ",0)", Me.Caption)
    
    
    Call ufgBorrow.SyncText(ufgBorrow.SelectionRow, gstrPatholCol_ȷ��״̬, BorrowSureState_NoSure, True)
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(False, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_�黹״̬) = "�ѹ黹")
    
    Call ufgMaterialDetail.SyncText(ufgMaterialDetail.SelectionRow, gstrPatholCol_�黹״̬, "�����", False)
    
'    Call RefreshStateInf(True, False)
End Sub

Private Sub DelArchivesBorrow(ByVal lngBorrowId As Long)
'ɾ������
    Dim strSql As String
    
    strSql = "Zl_�������_ɾ������(" & lngBorrowId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub


Private Sub ufgBackHistory_OnColFormartChange()
'������Ĺ黹�б��������
On Error GoTo errHandle
    zlDatabase.SetPara "���Ĺ黹�б�����", ufgBackHistory.GetColsString(ufgBackHistory), glngSys, G_LNG_PATHOLBORROW_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgBorrow_OnColFormartChange()
'��������б��������
On Error GoTo errHandle
    zlDatabase.SetPara "�����б�����", ufgBorrow.GetColsString(ufgBorrow), glngSys, G_LNG_PATHOLBORROW_NUM
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
    If CDate(ufgBorrow.Text(Row, gstrPatholCol_�黹����)) < mdtServicesTime And ufgBorrow.Text(Row, gstrPatholCol_�黹״̬) <> "�ѹ黹" Then
        ufgBorrow.DataGrid.Cell(flexcpPicture, Row, ufgBorrow.GetColIndex(gstrPatholCol_���ĺ�)) = picTimeOut
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
    
    '���������ϸ
    Call LoadBorrowDetail(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    '����黹��ʷ
    Call LoadBorrowReturnHistory(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    Call ReadBorrowInf(ufgBorrow.SelectionRow)
    
    Call ConfigBorrowModifyState(ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_ȷ��״̬) = BorrowSureState_Sure, ufgBorrow.Text(ufgBorrow.SelectionRow, gstrPatholCol_�黹״̬) = "�ѹ黹")
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMaterialDetail_OnColFormartChange()
'���������ϸ�б��������
On Error GoTo errHandle
    zlDatabase.SetPara "������ϸ�б�����", ufgMaterialDetail.GetColsString(ufgMaterialDetail), glngSys, G_LNG_PATHOLBORROW_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgMaterialDetail_OnColsNameReSet()
On Error GoTo errHandle

    '���������ϸ
    Call LoadBorrowDetail(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
