VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholLoseStation 
   Caption         =   "���������ʧ����վ"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   -450
   ClientWidth     =   13530
   Icon            =   "frmPatholLoseStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   13530
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   6360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   13335
      TabIndex        =   6
      Top             =   1680
      Width           =   13335
      Begin VB.TextBox txtStudyDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtStudyType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtStudyItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "������ڣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11160
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "������ͣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9240
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "�����Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "�� �䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "�� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "�� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2040
      Top             =   3960
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
            Picture         =   "frmPatholLoseStation.frx":1042
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":1D1C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":29F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":36D0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":43AA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":5084
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":5D5E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":6A38
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":7712
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":83EC
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":90C6
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenus 
      Left            =   2880
      Top             =   3960
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
            Picture         =   "frmPatholLoseStation.frx":A118
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":A46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":A7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":AB3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":AE90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B1E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B534
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B886
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":BBD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":BF2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C920
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":CC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":CFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D316
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D668
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D9BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":DD0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E3B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E702
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":EA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":EDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F44A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":FAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":FE40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame framQuery 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13335
      Begin VB.OptionButton Option2 
         Caption         =   "��ʧ���ڲ�ѯ"
         Height          =   375
         Left            =   100
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����Ų�ѯ"
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3240
         TabIndex        =   24
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   56623107
         CurrentDate     =   40928
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   1560
         TabIndex        =   22
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   56623107
         CurrentDate     =   40928
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "�ؼ����"
         Height          =   180
         Index           =   2
         Left            =   10920
         TabIndex        =   21
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "��Ƭ����"
         Height          =   180
         Index           =   1
         Left            =   9840
         TabIndex        =   20
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "�������"
         Height          =   180
         Index           =   0
         Left            =   8760
         TabIndex        =   19
         Top             =   330
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "�� ѯ(&Q)"
         Height          =   400
         Left            =   7440
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtPatholNo 
         Height          =   300
         Left            =   6000
         TabIndex        =   4
         Top             =   280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "��"
         Height          =   255
         Left            =   2985
         TabIndex        =   23
         Top             =   330
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ԥ��"
            Key             =   "tbn_PreviewList"
            Object.Tag             =   "Ԥ��"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ӡ"
            Key             =   "tbn_PrintList"
            Object.Tag             =   "��ӡ"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "������ʧ"
            Key             =   "tbn_NewLose"
            Object.Tag             =   "������ʧ"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����һ�"
            Key             =   "tbn_FindLose"
            Object.Tag             =   "�����һ�"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "tbn_Help"
            Object.Tag             =   "����"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "tbn_Exit"
            Object.Tag             =   "�˳�"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin zl9PACSWork.ucFlexGrid ufgLose 
      Height          =   6570
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11589
      DefaultCols     =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7575
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPatholLoseStation.frx":10192
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "��ʧ����������"
            TextSave        =   "��ʧ����������"
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14237
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnu_File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "��ӡ����(&S)"
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
         Caption         =   "�����Excel(&E)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "�˳�(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnu_MaterialLose 
         Caption         =   "������ʧ(&E)"
      End
      Begin VB.Menu mnu_MaterialFind 
         Caption         =   "�����һ�(&F)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnu_ToolsBar 
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
      Begin VB.Menu mnu_Split3 
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
      Begin VB.Menu mnu_HelpMain 
         Caption         =   "��������(&H)"
      End
      Begin VB.Menu mnu_WebZl 
         Caption         =   "WEB�ϵ�����(&W)"
         Begin VB.Menu mnu_HomePage 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnu_BBS 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnu_back 
            Caption         =   "���ͷ���(&K)"
         End
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "����...(&A)"
      End
   End
End
Attribute VB_Name = "frmPatholLoseStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



#Const DebugState = False


'Ϊ�˵�������Ӧ��ͼ��
Private Const MF_BITMAP = &H400&

Private Enum TQueryWay
    qwLoseDate = 0
    qwPatholNum = 1
End Enum


Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Private mstrPrivs As String
Private mblnMoved As Boolean

Private mqwQueryWay As TQueryWay
Private mstrCurSelectPatholNum As String

Private Sub InitMenuIcoConfig()
'��ʼ���˵�ͼ����ʾ
On Error Resume Next
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim hSubSubMenu As Long
    
    hMenu = GetMenu(Me.hWnd)
    
    '���õ�һ��˵�(�ļ�)
    hSubMenu = GetSubMenu(hMenu, 0) 'ȡ�õ�һ��˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(3).Picture, imgMenus.ListImages(3).Picture) '��ӡ����
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '��ӡԤ��
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '��ӡ
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '����Excel
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '�˳�
    

    '���õڶ���˵����༭��
    hSubMenu = GetSubMenu(hMenu, 1) 'ȡ�õڶ���˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '������ʧ
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '�����һ�
    
    
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


Private Sub ConfigPopedomFace()
'����Ȩ�����ý��棬������߱�Ȩ��ʱ�������ض�Ӧ���ܰ�ť
    Dim i As Long
    
    mnu_MaterialFind.Visible = CheckPopedom(mstrPrivs, "�����һ�")
    
    mnu_MaterialLose.Visible = CheckPopedom(mstrPrivs, "������ʧ")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_NewLose") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "������ʧ")
            
        End If
    Next i
End Sub




Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim curDate As Date
'    #If DebugState = True Then
'        Call InitDebugObject(1294, Me, "zlhis", "HIS")
'    #End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    mstrPrivs = gstrPrivs
    
    Call InitMenuIcoConfig
    
    Call InitLoseList
    
    Call RefreshStateInf
    
    curDate = zlDatabase.Currentdate
    
    dtpStart.value = Format(DateAdd("m", -1, curDate), "yyyy-mm-dd 00:00:00")
    dtpEnd.value = Format(curDate, "yyyy-mm-dd 23:59:59")
    
    mqwQueryWay = qwPatholNum
    mstrCurSelectPatholNum = ""
Exit Sub
ErrHandle:
If ErrCenter() = 1 Then Resume
End Sub



Private Sub RefreshStateInf()
'ˢ�²�����ʧ����
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select sum(��ʧ����) as ����ֵ from ������ʧ��Ϣ"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsData.RecordCount > 0 Then
        stbThis.Panels(2).Text = "��ʧ����������" & Nvl(rsData!����ֵ)
    End If
End Sub


Private Sub chkMaterial_Click(Index As Integer)
'���˲�ͬ���Ĳ���
On Error GoTo ErrHandle
    Dim strFilter As String
    
    strFilter = ""
    If chkMaterial(0).value <> 0 Then
        strFilter = " �������='����'"
    End If
    
    If chkMaterial(1).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & " �������='��Ƭ'"
    End If
    
    If chkMaterial(2).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & " �������='����' or �������='����' or �������='��Ⱦ' "
    End If
    
    If ufgLose.AdoData Is Nothing Then Exit Sub
    
    ufgLose.AdoData.Filter = strFilter
    
    Call ufgLose.RefreshData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Resize()
On Error Resume Next
    framQuery.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    framQuery.Left = 120
    framQuery.Width = Me.ScaleWidth - 240
    
    picInfo.Top = framQuery.Top + framQuery.Height
    picInfo.Left = 120
    picInfo.Width = Me.ScaleWidth - 240
    
    ufgLose.Top = picInfo.Top + picInfo.Height
    ufgLose.Left = 120
    ufgLose.Width = Me.ScaleWidth - 240
    ufgLose.Height = Me.ScaleHeight - framQuery.Height - picInfo.Height - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
err.Clear
End Sub


Private Sub cmdQuery_Click()
'��ѯ����
On Error GoTo ErrHandle
    mblnMoved = MovedByDate(dtpStart.value)
    
    If txtPatholNo.Enabled Then
        mqwQueryWay = qwPatholNum
        
        Call QueryStudyInf(txtPatholNo.Text)
        
        Call ufgLose.ClearListData
        
        If txtPatholNo.Text = "" Then Exit Sub
    Else
        mqwQueryWay = qwLoseDate
        
        Call QueryStudyInf("")
    End If
    
    Call QueryPatholMaterialData(txtPatholNo.Text, Format(dtpStart.value, "yyyy-mm-dd 23:59:59"), Format(dtpEnd.value, "yyyy-mm-dd 23:59:59"))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DTPEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub dtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        dtpEnd.SetFocus
    End If
    
End Sub

Private Sub QueryStudyInf(ByVal strPatholNum As String)
'��ѯ��������Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If Trim(strPatholNum) = "" Then Exit Sub
    
    strSQL = "select b.����, b.�Ա�,b.����,b.ҽ������,decode(a.�������,0,'����',1,'����',2,'ϸ��',3,'����',4,'ʬ��','����ʯ��') as �������,a.����ʱ�� " & _
            " from ����ҽ����¼ b , ��������Ϣ a where a.ҽ��ID=b.Id and a.�����=[1]"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatholNum)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtName.Text = Nvl(rsData!����)
    txtSex.Text = Nvl(rsData!�Ա�)
    txtAge.Text = Nvl(rsData!����)
    txtStudyItem.Text = Nvl(rsData!ҽ������)
    txtStudyType.Text = Nvl(rsData!�������)
    txtStudyDate.Text = Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd")
End Sub


'
'Private Sub QueryPatholMaterialDataByDate(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
''��ѯ�������
'    Dim strSql As String
'    Dim strLinkTable As String
'
'
'    Call ufgLose.ClearDataList
'
'    'ͳ����ʧ�Ĳ���������ͳ�ƽ�������ʱ��ֻ��ͳ��δ�黹�Ľ������������ֹ滮������ʧ�Ĳ��Ͻ�������ʧ���������ֵ���ʧ�����У�
'    strLinkTable = " (select nvl(sum(��ʧ����),0) as ��ʧ����, �鵵ID " & _
'                    " from ������ʧ��Ϣ Where ��ʧ���� between [1] and [2] group by �鵵ID ) x, " & _
'                    " (select (nvl(sum(��������), 0) - nvl(sum(�黹����), 0)) as �ѽ�����, a.�鵵ID " & _
'                    " from ������Ĺ��� a where a.�黹״̬=0 and  a.�鵵ID in(select �鵵ID from ������ʧ��Ϣ where ��ʧ���� between [1] and [2]) " & _
'                    " group by a.�鵵ID" & ") y"
'
'
'
'    strSql = "select distinct d.�������, d.�����, a.id, c.���, c.�걾����, c.ȡ��λ��, '����' as �������, " & _
'            " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
'            " (c.������ - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0) ) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " f.��������, '����:' || f.�������� || ' ���:' || f.������� || ' ����:' || f.�������� as ���λ��, f.��ϸ��ַ " & _
'            " from ��������Ϣ f, ��������Ϣ d, ����ȡ����Ϣ c, ����鵵��Ϣ a, ������ʧ��Ϣ g," & strLinkTable & _
'            " where f.id=a.����id and d.����ҽ��id=c.����ҽ��id and c.�Ŀ�id=a.�Ŀ�id and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.ID=g.�鵵ID and g.��ʧ���� between [1] and [2] " & _
'        " Union All " & _
'            " select distinct d.�������, d.�����, a.id, c.���, c.�걾����, c.ȡ��λ��, '��Ƭ' as �������, " & _
'            " decode(b.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
'            " (b.��Ƭ�� - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
'            " from ��������Ϣ e, ��������Ϣ d, ����ȡ����Ϣ c, ������Ƭ��Ϣ b, ����鵵��Ϣ a, ������ʧ��Ϣ f," & strLinkTable & _
'            " where e.id = a.����id and  d.����ҽ��id=c.����ҽ��id and c.�Ŀ�id=b.�Ŀ�id and b.id=a.��Ƭid and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.ID=f.�鵵ID and f.��ʧ���� between [1] and [2] " & _
'        " Union All " & _
'            " select distinct d.�������, d.�����,  a.id, c.���, c.�걾����, c.ȡ��λ��, decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
'            " decode(b.�ؼ�ϸĿ,0,decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� || decode(b.��������,-1,'-��',0,'','-��' || b.��������) || ')' as ������ϸ, " & _
'            " (1 - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
'            " from ��������Ϣ e, ��������Ϣ f, ��������Ϣ d, ����ȡ����Ϣ c, �����ؼ���Ϣ b, ����鵵��Ϣ a, ������ʧ��Ϣ g, " & strLinkTable & _
'            " where e.id = a.����id and f.����ID=b.����ID and d.����ҽ��id=c.����ҽ��id and c.�Ŀ�id=b.�Ŀ�id and b.id=a.�ؼ�id and  a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) " & _
'            " and a.ID=g.�鵵ID and  g.��ʧ���� between [1] and [2] "
'
'    If mblnMoved Then
'        strSql = strSql & " Union All " & GetMovedDataSql(strSql)
'    End If
'
'    strSql = "select /*+RULE*/ * from ( " & strSql & ") order by �������, �����,���,������ϸ,���״̬"
'
'    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'    Call ufgLose.RefreshData
'
'
'    If ufgLose.AdoData.RecordCount <= 0 Then
'        Call MsgBoxD(Me, "δ��ѯ��������ݡ�", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
'End Sub



'Private Sub QueryPatholMaterialDataByPatholNum(ByVal strPatholNum As String)
''��ѯ�������
'    Dim strSql As String
'    Dim strLinkTable As String
'
'
'    Call ufgLose.ClearDataList
'
'    If Trim(txtPatholNo.Text) = "" Then Exit Sub
'
'    'ͳ����ʧ�Ĳ���������ͳ�ƽ�������ʱ��ֻ��ͳ��δ�黹�Ľ������������ֹ滮������ʧ�Ĳ��Ͻ�������ʧ���������ֵ���ʧ�����У�
'    strLinkTable = " (select nvl(sum(��ʧ����),0) as ��ʧ����, �鵵ID " & _
'                    " from ������ʧ��Ϣ a, ����鵵��Ϣ b, ��������Ϣ d Where a.�鵵ID = b.ID And b.����ҽ��id = d.����ҽ��id " & _
'                    " and d.�����=[1] group by �鵵ID ) x, " & _
'                    " (select (nvl(sum(��������), 0) - nvl(sum(�黹����), 0)) as �ѽ�����, �鵵ID " & _
'                    " from ������Ĺ��� a, ����鵵��Ϣ b, ��������Ϣ d where a.�鵵ID = b.ID And b.����ҽ��id = d.����ҽ��id " & _
'                    " and a.�黹״̬=0  and d.�����=[1] group by �鵵ID" & ") y"
'
'
'
'    strSql = "select * from (select d.�������, d.�����, a.id, c.���, c.�걾����, c.ȡ��λ��, '����' as �������, " & _
'            " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
'            " (c.������ - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0) ) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " f.��������, '����:' || f.�������� || ' ���:' || f.������� || ' ����:' || f.�������� as ���λ��, f.��ϸ��ַ " & _
'            " from ����鵵��Ϣ a, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ f, " & strLinkTable & _
'            " where a.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.����id=f.id  and d.�����=[1] " & _
'        " Union All " & _
'            " select d.�������, d.�����, a.id, c.���, c.�걾����, c.ȡ��λ��, '��Ƭ' as �������, " & _
'            " decode(b.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
'            " (b.��Ƭ�� - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
'            " from ����鵵��Ϣ a, ������Ƭ��Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ e, " & strLinkTable & _
'            " where a.��Ƭid=b.id and b.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.����id=e.id  and d.�����=[1] " & _
'        " Union All " & _
'            " select d.�������, d.�����, a.id, c.���, c.�걾����, c.ȡ��λ��, decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
'            " decode(b.�ؼ�ϸĿ,0,decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� || decode(b.��������,-1,'-��',0,'','-��' || b.��������) || ')' as ������ϸ, " & _
'            " (1 - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ڵ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬,a.����״̬, " & _
'            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
'            " from ����鵵��Ϣ a, �����ؼ���Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ e, ��������Ϣ f, " & strLinkTable & _
'            " where a.�ؼ�id=b.id and b.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) " & _
'            " and a.����id=e.id and b.����ID=f.����ID and d.�����=[1] " & _
'        ") order by �������, ���,������ϸ,���״̬"
'
'
'    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatholNum)
'
'    Call ufgLose.RefreshData
'
'
'    If ufgLose.AdoData.RecordCount <= 0 Then
'        Call MsgBoxD(Me, "δ��ѯ��������ݡ�", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
'End Sub


Private Sub QueryPatholMaterialData(ByVal strPatholNum As String, ByVal dtStartDate As Date, ByVal dtEndDate As Date)
'��ѯ�������
    Dim strSQL As String
    Dim strFilter As String
    
    Dim strSqlMaterial As String
    Dim strSqlSlices As String
    Dim strSqlSpecial As String
    Dim strSqlMaterialCount As String
    Dim strSqlLoseCount As String
    
    
    Call ufgLose.ClearListData
    
    If strPatholNum <> "" Then
        strFilter = " and b.�����=[1] "
        
        strSqlMaterialCount = " select �鵵ID, sum(nvl(��������,0)) - sum(nvl(�黹����,0)) as �ѽ����� " & _
                                " from ��������Ϣ h, ����鵵��Ϣ i, ������Ĺ��� j " & _
                                " where h.����ҽ��id=i.����ҽ��id and i.id = j.�鵵id and h.�����=[1] and j.�黹״̬=0 " & _
                                " and not exists(select 1 from ������ʧ��Ϣ where ����ID=j.����id and �鵵ID=j.�鵵id) " & _
                                " group by �鵵ID"
                                
        strSqlLoseCount = " select �鵵ID, sum(nvl(��ʧ����,0)) as ����ʧ���� " & _
                            " from ��������Ϣ h, ����鵵��Ϣ i, ������ʧ��Ϣ j " & _
                            " where h.����ҽ��id=i.����ҽ��id and i.id = j.�鵵id and  h.�����=[1] " & _
                            " group by �鵵ID"
    Else
        strFilter = " and e.�鵵ID in(select �鵵ID from ������ʧ��Ϣ��where ��ʧ���� between [2] and [3]) "
        
        strSqlMaterialCount = " select j.�鵵ID, sum(nvl(j.��������,0)) - sum(nvl(j.�黹����,0)) as �ѽ����� " & _
                                " from ��������Ϣ h, ����鵵��Ϣ i, ������Ĺ��� j " & _
                                " Where h.����ҽ��id = i.����ҽ��id And i.ID = j.�鵵id and j.�黹״̬=0 " & _
                                " and j.�鵵id in(select �鵵id from ������ʧ��Ϣ where ����id is null " & _
                                " and  ��ʧ���� between [2] and [3]) " & _
                                " group by j.�鵵ID "
                                
        strSqlLoseCount = " select j.�鵵ID, sum(nvl(j.��ʧ����,0)) as ����ʧ���� " & _
                            " from ��������Ϣ h, ����鵵��Ϣ i, ������ʧ��Ϣ j " & _
                            " Where h.����ҽ��id = i.����ҽ��id And i.ID = j.�鵵id " & _
                            " and j.�鵵ID in (select �鵵ID  from ������ʧ��Ϣ where ��ʧ���� between [2] and [3]) " & _
                            " group by j.�鵵ID  "
    End If
        
    'ͳ����ʧ�Ĳ���������ͳ�ƽ�������ʱ��ֻ��ͳ��δ�黹�Ľ������������ֹ滮������ʧ�Ĳ��Ͻ�������ʧ���������ֵ���ʧ�����У�
    
    
    strSqlMaterial = " select c.id, b.�������, '����' as �������, c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��,a.������ as ����, e.��ʧ����, " & _
                    " decode( e.����id, null,'�ڲ���ʧ', '������ʧ') as ��ʧԭ��,d.��������, '����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, " & _
                    " case when a.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ " & _
                    " from ����ȡ����Ϣ a, ��������Ϣ b, ����鵵��Ϣ c, ��������Ϣ d, ������ʧ��Ϣ e " & _
                    " where a.����ҽ��id=b.����ҽ��id and a.�Ŀ�id=c.�Ŀ�id and c.����id=d.id and c.id=e.�鵵id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.�������, '����' as �������,  c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��, a.������ as ����, 0 as ��ʧ����, '����ʧ' as ��ʧԭ��,d.��������, " & _
                    " '����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, case when a.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ " & _
                    " from ����ȡ����Ϣ a, ��������Ϣ b,  ����鵵��Ϣ c, ��������Ϣ d " & _
                    " Where a.����ҽ��id = b.����ҽ��id And a.�Ŀ�id = c.�Ŀ�id And c.����id = d.ID " & _
                    " and not exists(select 1 from ������ʧ��Ϣ where �鵵ID=c.id) " & strFilter)
    
    strSqlSlices = " select c.id,  b.�������, '��Ƭ' as �������,c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��, x.��Ƭ�� as ����, e.��ʧ����, " & _
                    " decode( e.����id, null,'�ڲ���ʧ', '������ʧ') as ��ʧԭ��,d.��������, '����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, " & _
                    " decode(x.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ " & _
                    " from ������Ƭ��Ϣ x, ����ȡ����Ϣ a, ��������Ϣ b, ����鵵��Ϣ c, ��������Ϣ d, ������ʧ��Ϣ e " & _
                    " where x.�Ŀ�id=a.�Ŀ�id and a.����ҽ��id=b.����ҽ��id and x.id=c.��Ƭid and c.����id=d.id and c.id=e.�鵵id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.�������, '��Ƭ' as �������, c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��,x.��Ƭ�� as ����, 0 as ��ʧ����, '����ʧ' as ��ʧԭ��,d.��������, " & _
                    " '����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, decode(x.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ " & _
                    " from ������Ƭ��Ϣ x, ����ȡ����Ϣ a, ��������Ϣ b, ����鵵��Ϣ c, ��������Ϣ d " & _
                    " Where X.�Ŀ�id = a.�Ŀ�id And a.����ҽ��id = b.����ҽ��id And X.ID = c.��Ƭid And c.����id = d.ID " & _
                    " and not exists(select 1 from ������ʧ��Ϣ where �鵵ID=c.id) " & strFilter)


    strSqlSpecial = " select c.id,  b.�������, decode(x.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������,c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��, 1 as ����, e.��ʧ����, " & _
                    " decode( e.����id, null,'�ڲ���ʧ', '������ʧ') as ��ʧԭ��,d.��������,'����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, " & _
                    " decode(x.�ؼ�ϸĿ,0,decode(x.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� " & _
                    " || decode(x.��������,-1,'-��',0,'','-��' || x.��������) || ')' as ������ϸ " & _
                    " from �����ؼ���Ϣ x, ����ȡ����Ϣ a, ��������Ϣ b, ����鵵��Ϣ c, ��������Ϣ d, ������ʧ��Ϣ e, ��������Ϣ f " & _
                    " where x.�Ŀ�id=a.�Ŀ�id and a.����ҽ��id=b.����ҽ��id and x.id=c.�ؼ�id and c.����id=d.id and c.id=e.�鵵id and x.����id=f.����id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.�������, decode(x.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, c.���״̬, b.�����,a.���,a.�걾����,a.ȡ��λ��,1 as ����,0 as ��ʧ����, '����ʧ' as ��ʧԭ��,d.��������, " & _
                    " '����:' || d.�������� || ' ���:' || d.������� || ' ����:' || d.�������� as ���λ��, " & _
                    " decode(x.�ؼ�ϸĿ,0,decode(x.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� " & _
                    " || decode(x.��������,-1,'-��',0,'','-��' || x.��������) || ')' as ������ϸ " & _
                    " from �����ؼ���Ϣ x, ����ȡ����Ϣ a, ��������Ϣ b, ����鵵��Ϣ c, ��������Ϣ d, ��������Ϣ f " & _
                    " Where X.�Ŀ�id = a.�Ŀ�id And a.����ҽ��id = b.����ҽ��id And X.ID = c.�ؼ�id And c.����id = d.ID And X.����id = f.����id " & _
                    " and not exists(select 1 from ������ʧ��Ϣ where �鵵ID=c.id) " & strFilter)
    
    strSQL = "select id,�������,�������,���״̬,�����,���, �걾����,ȡ��λ��,������ϸ, decode(sum(nvl(��ʧ����,0)), 0, '����ʧ', ��ʧԭ��) as ��ʧԭ��,��������,���λ��, " & _
                " (nvl(����,0) - nvl(�ѽ�����, 0) - nvl(����ʧ����, 0)) as �ڵ�����, sum(nvl(��ʧ����,0)) as ��ʧ���� " & _
                " from( " & strSqlMaterial & " union all " & strSqlSlices & " union all " & strSqlSpecial & ") u, (" & _
                strSqlMaterialCount & ")v, (" & strSqlLoseCount & ") w" & _
                " where u.id =v.�鵵ID(+) and u.id = w.�鵵ID(+) " & _
                " group by id,�������,�������,���״̬,�����,���, �걾����,ȡ��λ��, ��ʧԭ��,��������,���λ��,������ϸ,����,�ѽ�����,����ʧ���� " & _
                IIf(strPatholNum = "", " having sum(nvl(��ʧ����,0)) > 0 ", "") & _
                " order by �������,�������,�����,���, ��ʧԭ�� "
    
    
    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatholNum, dtStartDate, dtEndDate)
                                                    
    Call ufgLose.RefreshData
                                                          

    If ufgLose.AdoData.RecordCount <= 0 Then
        mstrCurSelectPatholNum = ""
        txtName.Text = ""
        txtSex.Text = ""
        txtAge.Text = ""
        txtStudyItem.Text = ""
        txtStudyType.Text = ""
        txtStudyDate.Text = ""
        Call MsgBoxD(Me, "δ��ѯ��������ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
End Sub


Private Sub InitLoseList()
'��ʼ�������б�
    Dim strTemp As String
    
        '��������
    ufgLose.GridRows = glngStandardRowCount
    '�����и�
    ufgLose.RowHeightMin = glngStandardRowHeight
    
    strTemp = zlDatabase.GetPara("��ʧ�б�����", glngSys, G_LNG_PATHOLLOSE_NUM, "")
    
    ufgLose.IsCopyMode = True
    ufgLose.IsKeepRows = False
    ufgLose.DefaultColNames = gstrMaterialLoseCols
    ufgLose.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialLoseCols)
    '��ֹ�Ҽ������б����ô���
    ufgLose.IsEjectConfig = False
    ufgLose.ColConvertFormat = gstrMaterialLoseConvertFormat
                                 
End Sub


Private Sub Execute_MaterialLose()
'������ʧ����
    Dim frmLose As frmPatholLoseEnreg
    
    If Not ufgLose.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ������ʧ����Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_���״̬) = "����ʧ" Then
        Call MsgBoxD(Me, "�ò�������ʧ�����ܽ�����ʧ����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
On Error GoTo errFree
        Set frmLose = New frmPatholLoseEnreg
        
        Call frmLose.ShowLoseWindow(ufgLose, Me)
        
        If frmLose.blnIsOk Then
            Call RefreshStateInf
        End If
        
errFree:
    Call Unload(frmLose)
    Set frmLose = Nothing
End Sub


Private Sub Execute_MaterialFind()
'�����һش���
    Dim frmFind As frmPatholLoseEnreg
    
    If Not ufgLose.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ������ʧ����Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_���״̬) = "�浵��" Then
        Call MsgBoxD(Me, "�ò��ϴ��ڴ浵�У����ܽ����һش���", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_��ʧԭ��) = "������ʧ" Then
        Call MsgBoxD(Me, "����Ĳ�������ʧ��ֻ��ͨ�����Ĺ黹�һ���ʧ�Ĳ��ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
On Error GoTo errFree
        Set frmFind = New frmPatholLoseEnreg
        
        Call frmFind.ShowFindWindow(ufgLose, Me)
        
        If frmFind.blnIsOk Then
            Call RefreshStateInf
        End If
        
errFree:
    Call Unload(frmFind)
    Set frmFind = Nothing
End Sub

Private Sub Execute_Help()
'����
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub mnu_About_Click()
'����
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_back_Click()
'���ͷ���
On Error GoTo ErrHandle
    Call zlMailTo(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_BBS_Click()
'������̳
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Exit_Click()
'�˳�
On Error GoTo ErrHandle
    Call Unload(Me)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ExportExcel_Click()
'����Excel
On Error GoTo ErrHandle
    Call MenuPrint(3)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Font_Click()
'����
On Error GoTo ErrHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgLose.DataGrid.Font.Bold
    diaFont.FontName = ufgLose.DataGrid.Font.Name
    diaFont.FontSize = ufgLose.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgLose.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgLose.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '�����б�
    ufgLose.DataGrid.Font.Bold = diaFont.FontBold
    ufgLose.DataGrid.Font.Name = diaFont.FontName
    ufgLose.DataGrid.Font.Size = diaFont.FontSize
    ufgLose.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgLose.DataGrid.Font.Underline = diaFont.FontUnderline
    
    
    Call ufgLose.DataGrid.Refresh
    
    ufgLose.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgLose.DataGrid.AutoSize(0, ufgLose.DataGrid.Rows - 1)
    
    ufgLose.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgLose.DataGrid.AutoSize(0, ufgLose.DataGrid.Rows - 1)
    
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HelpMain_Click()
'����
On Error GoTo ErrHandle
    Call Execute_Help
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HomePage_Click()
'������ҳ
On Error GoTo ErrHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MaterialFind_Click()
'�����һ�
On Error GoTo ErrHandle
    Call Execute_MaterialFind
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MaterialLose_Click()
'������ʧ
On Error GoTo ErrHandle
    Call Execute_MaterialLose
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Preview_Click()
'Ԥ�������б�
On Error GoTo ErrHandle
    Call MenuPrint(0)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Print_Click()
'Ԥ�������б�
On Error GoTo ErrHandle
    Call MenuPrint(1)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintConfig_Click()
'��ӡ����
On Error GoTo ErrHandle
    Call zlPrintSet
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StandardButton_Click()
On Error GoTo ErrHandle
    Dim intCount As Long
    Me.mnu_StandardButton.Checked = Not Me.mnu_StandardButton.Checked
    Me.tbrTools.Visible = Me.mnu_StandardButton.Checked
    
    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If

    Me.tbrTools.Refresh
    
    Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StateBar_Click()
On Error GoTo ErrHandle
    Me.mnu_StateBar.Checked = Not Me.mnu_StateBar.Checked
    Me.stbThis.Visible = Me.mnu_StateBar.Checked
    
    Call Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_WordLabel_Click()
On Error GoTo ErrHandle
    Dim intCount As Long
    
    Me.mnu_WordLabel.Checked = Not Me.mnu_WordLabel.Checked

    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If
    
    Me.tbrTools.Refresh
    
    Call Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Option1_Click()
On Error Resume Next
    txtPatholNo.Enabled = True
    txtPatholNo.BackColor = &H80000005
    
    dtpStart.Enabled = False
    dtpEnd.Enabled = False
    
    err.Clear
End Sub

Private Sub Option2_Click()
On Error Resume Next
    txtPatholNo.Enabled = False
    txtPatholNo.Text = ""
    txtPatholNo.BackColor = &H8000000F
    
    dtpStart.Enabled = True
    dtpEnd.Enabled = True
    
    err.Clear
End Sub

Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle
    
    Select Case UCase(Button.Key)
        Case UCase("tbn_PreviewList")   'Ԥ��
            Call MenuPrint(0)
            
        Case UCase("tbn_PreviewPrint")  '��ӡ
            Call MenuPrint(1)
            
        Case UCase("tbn_NewLose")   '������ʧ
            Call Execute_MaterialLose
    
        Case UCase("tbn_FindLose")   '�����һ�
            Call Execute_MaterialFind
                        
        Case UCase("tbn_Help")  '����
            Call Execute_Help
            
        Case UCase("tbn_Exit")  '�Ƴ���������ģ��
            Call Unload(Me)
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��ӡԤ��, 0��������ѡ��Ի���1Ԥ����2��ӡ��3����Excel
    '������    �����ʽ
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgLose.DataGrid
    
    objPrint.Title = "�������״̬�嵥"

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


Private Sub txtPatholNo_KeyPress(KeyAscii As Integer)
 '�س�ִ�в�ѯ
 On Error GoTo ErrHandle
    If KeyAscii = 13 Then
        mblnMoved = MovedByDate(dtpStart.value)
        mqwQueryWay = qwPatholNum
        
        Call QueryStudyInf(txtPatholNo.Text)
        Call ufgLose.ClearListData
        
        If txtPatholNo.Text = "" Then Exit Sub
        
        Call QueryPatholMaterialData(txtPatholNo.Text, Format(dtpStart.value, "yyyy-mm-dd 23:59:59"), Format(dtpEnd.value, "yyyy-mm-dd 23:59:59"))
    End If

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgLose_OnColFormartChange()
On Error GoTo ErrHandle
    zlDatabase.SetPara "��ʧ�б�����", ufgLose.GetColsString(ufgLose), glngSys, G_LNG_PATHOLLOSE_NUM
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgLose_OnSelChange()
On Error GoTo ErrHandle
    '��ѯ������Ϣ
    
    If mqwQueryWay = qwPatholNum Then Exit Sub
    If Not ufgLose.IsSelectionRow Then Exit Sub
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_�����) = "" Then Exit Sub
    If mstrCurSelectPatholNum = ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_�����) Then Exit Sub
    
    mstrCurSelectPatholNum = ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_�����)
    
    Call QueryStudyInf(mstrCurSelectPatholNum)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
