VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholArchivesStation 
   Caption         =   "����鵵����վ"
   ClientHeight    =   8895
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14760
   Icon            =   "frmPatholArchivesStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14760
   StartUpPosition =   3  '����ȱʡ
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   7620
      Left            =   4455
      TabIndex        =   29
      Top             =   840
      Width           =   100
      _ExtentX        =   185
      _ExtentY        =   13441
      BackColor       =   -2147483633
      SplitWidth      =   100
      SplitLevel      =   3
      SyncParentHeight=   0   'False
      AllowPaintOtherSpliter=   -1  'True
      Control1Name    =   "Picture1"
      Control2Name    =   "Picture2"
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   4320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgMenus 
      Left            =   4680
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":179A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":1AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2512
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":35AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":3C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":3FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":42F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4646
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4998
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":503C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":538E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":56E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":5A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":5D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":60D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6428
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":677A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":7170
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":74C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":819C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":8E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":9B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":A82A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":B504
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":C1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":CEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":DB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":E86C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ǩԤ��"
            Key             =   "tbn_LabView"
            Object.Tag             =   "��ǩԤ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ǩ��ӡ"
            Key             =   "tbn_LabPrint"
            Object.Tag             =   "��ǩ��ӡ"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "tbn_NewArchives"
            Object.Tag             =   "��������"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ������"
            Key             =   "tbn_DelArchives"
            Object.Tag             =   "ɾ������"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���µ���"
            Key             =   "tbn_UpdateArchives"
            Object.Tag             =   "���µ���"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѯ����"
            Key             =   "tbn_QueryArchives"
            Object.Tag             =   "��ѯ����"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����鵵"
            Key             =   "tbn_EnterArchives"
            Object.Tag             =   "�����鵵"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����鵵"
            Key             =   "tbn_CancelArchives"
            Object.Tag             =   "�����鵵"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "tbn_Help"
            Object.Tag             =   "����"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "tbn_Exit"
            Object.Tag             =   "�˳�"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   4555
      ScaleHeight     =   7620
      ScaleWidth      =   10200
      TabIndex        =   1
      Top             =   840
      Width           =   10205
      Begin VB.TextBox txtNumberInf 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "��ǰ����������0   �ڵ�������0   �ѽ�������0   ��ʧ������0   "
         Top             =   90
         Width           =   5415
      End
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame framArchivesDetail 
         Height          =   6735
         Left            =   1320
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdFilter 
            Caption         =   "�� ��(&L)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "Ԥ ��(&W)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "ɾ ��(&D)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "�� ӡ(&P)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRead 
            Caption         =   "��ȡ��������(&R)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   6240
            Width           =   1695
         End
         Begin zl9PACSWork.ucFlexGrid ufgArchivesDetail 
            Height          =   5895
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   10398
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.Frame framEnterArchives 
         Height          =   7095
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   9735
         Begin VB.CommandButton cmdEnterArchives 
            Caption         =   "�����뵵(&I)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   5760
            TabIndex        =   42
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CheckBox chkTeShu 
            Caption         =   "�ؼ����"
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
            Left            =   4680
            TabIndex        =   25
            Top             =   6360
            Width           =   1215
         End
         Begin VB.CheckBox chkSlices 
            Caption         =   "��Ƭ����"
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
            Left            =   3600
            TabIndex        =   24
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CheckBox chkWaxStone 
            Caption         =   "�������"
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
            Left            =   2520
            TabIndex        =   23
            Top             =   6360
            Width           =   1215
         End
         Begin VB.CheckBox chkNotEnter 
            Caption         =   "��δ�뵵"
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
            Left            =   1200
            TabIndex        =   22
            Top             =   6360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkComplete 
            Caption         =   "������"
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
            TabIndex        =   21
            Top             =   6360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.ComboBox cbxRequestDetail 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F546
            Left            =   6120
            List            =   "frmPatholArchivesStation.frx":F548
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cbxRequestType 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F54A
            Left            =   3600
            List            =   "frmPatholArchivesStation.frx":F54C
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cbxStudyType 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F54E
            Left            =   1080
            List            =   "frmPatholArchivesStation.frx":F550
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   840
            Width           =   1455
         End
         Begin VB.Frame framQuery 
            Height          =   735
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   9735
            Begin VB.ComboBox cbxQueryType 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmPatholArchivesStation.frx":F552
               Left            =   120
               List            =   "frmPatholArchivesStation.frx":F55C
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdQuery 
               Caption         =   "��ѯ(&Q)"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   7680
               TabIndex        =   14
               Top             =   180
               Width           =   975
            End
            Begin VB.TextBox txtEndPatholNum 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6600
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtStartPatholNum 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   5400
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin MSComCtl2.DTPicker dtpStartDate 
               Height          =   330
               Left            =   1320
               TabIndex        =   7
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd 00:00:00"
               Format          =   114032643
               CurrentDate     =   40884
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   330
               Left            =   3015
               TabIndex        =   9
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd 23:59:59"
               Format          =   114032643
               CurrentDate     =   40884
            End
            Begin VB.Label Label3 
               Caption         =   "��"
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
               Left            =   6390
               TabIndex        =   12
               Top             =   300
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "����ţ�"
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
               Left            =   4680
               TabIndex        =   10
               Top             =   300
               Width           =   735
            End
            Begin VB.Label labTo 
               Caption         =   "��"
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
               Left            =   2805
               TabIndex        =   8
               Top             =   300
               Width           =   255
            End
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialQuery 
            Height          =   4935
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8705
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Line lineSplit2 
            BorderColor     =   &H00C0C0C0&
            X1              =   2400
            X2              =   2400
            Y1              =   6360
            Y2              =   6600
         End
         Begin VB.Label labRequestDetail 
            Caption         =   "���ϸĿ��"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "�����̣�"
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
            Left            =   2640
            TabIndex        =   17
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "������ͣ�"
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
            TabIndex        =   15
            Top             =   880
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   661
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   0
      ScaleHeight     =   7620
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   120
         ScaleHeight     =   7335
         ScaleWidth      =   4335
         TabIndex        =   31
         Top             =   120
         Width           =   4335
         Begin zl9PacsControl.ucSplitter ucSplitter2 
            Height          =   100
            Left            =   0
            TabIndex        =   32
            Top             =   3930
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   185
            BackColor       =   -2147483633
            MousePointer    =   7
            SplitWidth      =   100
            SplitType       =   0
            SplitLevel      =   3
            Control1Name    =   "ufgArchives"
            Control2Name    =   "rtbDetail"
         End
         Begin RichTextLib.RichTextBox rtbDetail 
            Height          =   3305
            Left            =   0
            TabIndex        =   33
            Top             =   4030
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   5821
            _Version        =   393217
            BackColor       =   16761024
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmPatholArchivesStation.frx":F574
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
         Begin zl9PACSWork.ucFlexGrid ufgArchives 
            Height          =   3930
            Left            =   0
            TabIndex        =   34
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   8535
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPatholArchivesStation.frx":F611
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "δ�鵵��������"
            TextSave        =   "δ�鵵��������"
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "δ�뵵��������"
            TextSave        =   "δ�뵵��������"
            Key             =   "sb_NoEnterWaxStone"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3176
            MinWidth        =   3176
            Text            =   "δ�뵵��Ƭ����"
            TextSave        =   "δ�뵵��Ƭ����"
            Key             =   "sb_NoEnterSlices"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "δ�뵵�ؼ�����"
            TextSave        =   "δ�뵵�ؼ�����"
            Key             =   "sb_NoEnterSpeEx"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6802
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
   Begin VB.Menu mnu_File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnu_ParameterConfig 
         Caption         =   "��������(&M)"
      End
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "��ӡ����(&C)"
      End
      Begin VB.Menu mnu_Split10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ArchivesClassCfg 
         Caption         =   "������������(&A)"
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ListPreview 
         Caption         =   "Ԥ ��(&V)"
      End
      Begin VB.Menu mnu_ListPrint 
         Caption         =   "�� ӡ(&P)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportExcel 
         Caption         =   "�����Excel(&E)"
      End
      Begin VB.Menu mnu_Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "�� ��(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnu_NewArchives 
         Caption         =   "��������(&N)"
      End
      Begin VB.Menu mnu_DelArchives 
         Caption         =   "ɾ������(&D)"
      End
      Begin VB.Menu mnu_UpdateArchives 
         Caption         =   "���µ���(&U)"
      End
      Begin VB.Menu mnu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EnterArchives 
         Caption         =   "�����鵵(&T)"
      End
      Begin VB.Menu mnu_CancelArchives 
         Caption         =   "�����鵵(&R)"
      End
      Begin VB.Menu mnu_Split6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_QueryArchives 
         Caption         =   "��ѯ����(&Q)"
      End
      Begin VB.Menu mnu_Split9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_LabPreview 
         Caption         =   "��ǩԤ��(&V)"
      End
      Begin VB.Menu mnu_LabPrint 
         Caption         =   "��ǩ��ӡ(&P)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnu_ToolsBar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnu_StandardBut 
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
         Caption         =   "�� ��(&F)"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "����(&T)"
      Visible         =   0   'False
      Begin VB.Menu mnu_Zoom 
         Caption         =   "�Ŵ�(&Z)"
      End
      Begin VB.Menu mnu_Calc 
         Caption         =   "������(&C)"
      End
   End
   Begin VB.Menu mnu_MainHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnu_Help 
         Caption         =   "��������(&H)"
      End
      Begin VB.Menu mnu_WebZL 
         Caption         =   "WEB�ϵ�����(&W)"
         Begin VB.Menu mnu_MainPage 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnu_BBS 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnu_Return 
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
Attribute VB_Name = "frmPatholArchivesStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False

Private Const ArchivesState_NoEnter As String = "δ�鵵"
Private Const ArchivesState_Enter As String = "�ѹ鵵"

'Ϊ�˵�������Ӧ��ͼ��
Private Const MF_BITMAP = &H400&


'������������ö��
Private Enum TArchivesMaterialType
    amtTable = 0
    amtMaterial = 1
    amdReport = 2
End Enum


Private mstrPrivs As String
Private mcurMaterialType As TArchivesMaterialType
Private mlngCurArchivesId As Long
Private mblnMoved As Boolean

Private mlngDefaultQueryDays As Long
Private mstrLabelReportName As String
Private mblnIsFormLoaded As Boolean



Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1



Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long 'ȡ�ô��ڵĲ˵����,hwnd�Ǵ��ڵľ��
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal npos As Long) As Long 'ȡ���Ӳ˵������nPos�ǲ˵���λ��
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal npos As Long, ByVal wFlags As Long, ByVal hBitUnchecked As Long, ByVal hBitChecked As Long) As Long







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
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(6).Picture, imgMenus.ListImages(6).Picture) '������������
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '��ӡԤ��
    Call SetMenuItemBitmaps(hSubMenu, 6, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '��ӡ
    Call SetMenuItemBitmaps(hSubMenu, 8, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '����Excel
    Call SetMenuItemBitmaps(hSubMenu, 10, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '�˳�
    

    '���õڶ���˵����༭��
    hSubMenu = GetSubMenu(hMenu, 1) 'ȡ�õڶ���˵����Ӳ˵����
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '��������
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(8).Picture, imgMenus.ListImages(8).Picture) 'ɾ������
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(9).Picture, imgMenus.ListImages(9).Picture) '���µ���
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '�����鵵
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(11).Picture, imgMenus.ListImages(11).Picture) '�����鵵
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(12).Picture, imgMenus.ListImages(12).Picture) '��ѯ����
    Call SetMenuItemBitmaps(hSubMenu, 9, MF_BITMAP, imgMenus.ListImages(1).Picture, imgMenus.ListImages(1).Picture) '��ӡԤ��
    Call SetMenuItemBitmaps(hSubMenu, 10, MF_BITMAP, imgMenus.ListImages(2).Picture, imgMenus.ListImages(2).Picture) '��ӡ
    
    
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


Private Sub RefreshStateInf(ByVal blnIsRefreshArchives As Boolean, ByVal blnIsRefreshMaterial As Boolean)
'ˢ��״̬��Ϣ����δ�鵵����������δ�鵵����������...
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If blnIsRefreshArchives Then
        'ˢ�µ�������
        strSql = "select /*+ Rule*/ count(1) as ����ֵ from ��������Ϣ where ����״̬=0"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(2).Text = "δ�鵵��������" & Nvl(rsData!����ֵ)
        End If
    End If
    
    If blnIsRefreshMaterial Then
        'ˢ����������
        strSql = "select /*+ Rule*/ count(1) as ����ֵ from ����ȡ����Ϣ a, ����ҽ������ b, ��������Ϣ c " & _
                " Where a.����ҽ��id = c.����ҽ��id And b.ҽ��ID = c.ҽ��ID And b.ִ�й��� = 6 And a.�鵵״̬ = 0 and a.ȷ��״̬=1 " & _
                " and a.ȡ��ʱ�� between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(3).Text = "δ�뵵��������" & Nvl(rsData!����ֵ)
        End If
    
    
    
        'ˢ����Ƭ����
        strSql = "select /*+ Rule*/ count(1) as ����ֵ from ������Ƭ��Ϣ a, ����ҽ������ b, ��������Ϣ c " & _
                " Where a.����ҽ��id = c.����ҽ��id And b.ҽ��ID = c.ҽ��ID And b.ִ�й��� = 6 And a.�鵵״̬ = 0 and a.��ǰ״̬=2 " & _
                " and a.��Ƭʱ�� between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(4).Text = "δ�뵵��Ƭ����" & Nvl(rsData!����ֵ)
        End If
        
        
        'ˢ���ؼ�����
        strSql = "select /*+ Rule*/ count(1) as ����ֵ from �����ؼ���Ϣ a, ����ҽ������ b, ��������Ϣ c " & _
                " Where a.����ҽ��id = c.����ҽ��id And b.ҽ��ID = c.ҽ��ID And b.ִ�й��� = 6 And a.�鵵״̬ = 0 and a.��ǰ״̬=2 " & _
                " and a.���ʱ�� between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(5).Text = "δ�뵵�ؼ�����" & Nvl(rsData!����ֵ)
        End If
    End If
    
End Sub


Private Sub QueryMaterialData()
'��ѯ��������
    Dim strSql As String
    Dim strPatholNumQuery As String
    Dim strFilterDate As String
    Dim strRequestFrom As String
    Dim lngCurArchivesId As Long
    
    lngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))


    strRequestFrom = ""
    strFilterDate = ""
    
    If cbxQueryType.Text = "����ʱ��" Then
        strFilterDate = " and a.����ʱ�� between [1] and [2] "
    Else
        strFilterDate = " and c.����ID=r.����ID and r.����ʱ�� between [1] and [2]"
        strRequestFrom = " ,����������Ϣ r "
    End If


    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(a.�����), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(a.�����), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(a.�����), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(a.�����),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(a.�����)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(a.�����) =upper([4]) "
    End If
    
    
    If mcurMaterialType <> amtMaterial Then
        '�Ȳ�ѯ��������ļ����Ϣ���ڵ�����ѯ�����Ϣ�����ڶԽ����еļ���״̬���ˣ�
        strSql = " select distinct a.����ҽ��ID, '' as �������, 0 as ���, 4 as ������Դ,  '4-' || a.����ҽ��ID as ��ԴID,a.�����, b.����,b.�Ա�,b.����, " & _
                " b.ҽ������ as �����Ŀ, a.�������, null as �Ŀ��, null as �걾����, null as ȡ��λ��, null as ������ϸ, null as ����״̬, " & _
                " null as ����, decode((select ����ID from ����鵵��Ϣ where ����ҽ��ID=a.����ҽ��ID and ����ID=[5]),[5],'�Ѵ���','δ�鵵') as ���״̬,  a.����ʱ��, 1 as �Ƿ�����, " & _
                " decode(r.��������, 0, '����',1,'��Ⱦ',2,'����',3,'����Ƭ',4,'��ȡ��','') as ��������, null as ����ϸĿ,c.ִ�й���  " & _
                " from ��������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ����������Ϣ r " & _
                " where a.ҽ��ID=b.id and a.ҽ��ID=c.ҽ��ID and a.����ҽ��id=r.����ҽ��ID  " & _
                IIf(cbxQueryType.Text = "����ʱ��", " and a.����ʱ�� between [1] and [2]", " and r.����ʱ�� between [1] and [2]") & strPatholNumQuery & _
                " Union All  " & _
                " select distinct a.����ҽ��ID, '' as �������, 0 as ���, 4 as ������Դ,  '4-' || a.����ҽ��ID as ��ԴID,a.�����, b.����,b.�Ա�,b.����,  " & _
                " b.ҽ������ as �����Ŀ, a.�������, null as �Ŀ��, null as �걾����, null as ȡ��λ��, null as ������ϸ, null as ����״̬, " & _
                " null as ����, decode((select ����ID from ����鵵��Ϣ where ����ҽ��ID=a.����ҽ��ID and ����ID=[5]), [5],'�Ѵ���','δ�鵵') as ���״̬,  a.����ʱ��, 0 as �Ƿ�����, null as ��������, null as ����ϸĿ,c.ִ�й���  " & _
                " from ��������Ϣ a, ����ҽ����¼ b, ����ҽ������ c " & IIf(cbxQueryType.Text = "����ʱ��", "", " , ����������Ϣ r") & _
                " where a.ҽ��ID=b.id and a.ҽ��ID=c.ҽ��ID  " & _
                IIf(cbxQueryType.Text = "����ʱ��", " and a.����ʱ�� between [1] and [2] ", " and a.����ҽ��id=r.����ҽ��ID and r.����ʱ�� between [1] and [2] ") & _
                strPatholNumQuery
    Else
        strFilterDate = strFilterDate & strPatholNumQuery
        
        strSql = "select  1 as ������Դ, '1-' || c.�Ŀ�ID as ��ԴID, a.����ҽ��id, a.�����, b.����, b.�Ա�, b.����,b.ҽ������ as �����Ŀ, e.ִ�й���, " & _
                " a.�������, c.���,c.ȡ��λ��,c.�걾����,to_number(c.������) as ����, a.����ʱ��, '����' as �������, '' ����ϸĿ," & _
                " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ��������, case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
                " case when d.���״̬ is null then 'δ�鵵' else decode(d.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') end as ���״̬ , k.��������, k.��ϸ��ַ, " & _
                " case when k.id is null then '' else '����:' || k.�������� || ' ���:' || k.������� || ' ����:' || k.�������� end as ���λ�� " & _
                " from  ��������Ϣ a, ����ҽ����¼ b, ����ȡ����Ϣ c, ��������Ϣ k, ����鵵��Ϣ d, ����ҽ������ e " & strRequestFrom & _
                " where a.ҽ��id=b.id and a.����ҽ��id=c.����ҽ��id and b.ID=e.ҽ��ID and k.id(+) = d.����ID and c.�Ŀ�id=d.�Ŀ�id(+) and c.ȷ��״̬=1 and c.������>0 and a.�������<>3 " & strFilterDate & _
                " Union All select 2 as ������Դ, '2-' || c.ID as ��ԴID, a.����ҽ��id, a.�����, b.����, b.�Ա�, b.����,b.ҽ������ as �����Ŀ, f.ִ�й���, " & _
                " a.�������, d.���,d.ȡ��λ��,d.�걾����,to_number(c.��Ƭ��) as ����, a.����ʱ��, '��Ƭ' as �������, " & _
                " decode(c.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') ����ϸĿ,case when c.����ID is null then '������Ƭ' else '����Ƭ' end as ��������, " & _
                " decode(c.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
                " case when e.���״̬ is null then 'δ�鵵' else decode(e.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') end as ���״̬, k.��������, k.��ϸ��ַ, " & _
                " case when k.id is null then '' else '����:' || k.�������� || ' ���:' || k.������� || ' ����:' || k.�������� end as ���λ��  " & _
                " from  ��������Ϣ a, ����ҽ����¼ b, ������Ƭ��Ϣ c, ����ȡ����Ϣ d, ��������Ϣ k, ����鵵��Ϣ e, ����ҽ������ f " & strRequestFrom & _
                " where a.ҽ��id=b.id and a.����ҽ��id = d.����ҽ��id and b.ID=f.ҽ��ID and d.�Ŀ�id=c.�Ŀ�id and k.id(+) = e.����ID and c.id=e.��Ƭid(+) and c.��ǰ״̬=2 " & strFilterDate & _
                " Union All select 3 as ������Դ, '3-' || c.ID as ��ԴID, a.����ҽ��id, a.�����, b.����, b.�Ա�, b.����,b.ҽ������ as �����Ŀ, g.ִ�й���, " & _
                " a.�������, d.���,d.ȡ��λ��,d.�걾����,1 as ����, a.����ʱ��, decode(c.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
                " decode(�ؼ�ϸĿ,1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') as ����ϸĿ, decode(c.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as ��������, " & _
                " decode(c.�ؼ�ϸĿ,0,decode(c.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� || decode(c.��������,-1,'-��',0,'','-��' || c.��������) || ')' as ������ϸ, " & _
                " case when e.���״̬ is null then 'δ�鵵' else decode(e.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') end as ���״̬, k.��������, k.��ϸ��ַ, " & _
                " case when k.id is null then '' else '����:' || k.�������� || ' ���:' || k.������� || ' ����:' || k.�������� end as ���λ�� " & _
                " from  ��������Ϣ a, ����ҽ����¼ b, �����ؼ���Ϣ c, ����ȡ����Ϣ d, ��������Ϣ k, ����鵵��Ϣ e, ��������Ϣ f,����ҽ������ g  " & strRequestFrom & _
                " where a.ҽ��id=b.id and a.����ҽ��id = d.����ҽ��id  and b.ID=g.ҽ��ID and d.�Ŀ�id=c.�Ŀ�id and k.id(+) = e.����ID and c.id=e.�ؼ�id(+) and c.����ID=f.����id and c.��ǰ״̬=2 " & strFilterDate
    End If
    
'    If mblnMoved Then
'        strSql = strSql & " union all " & GetMovedDataSql(strSql)
'    End If
    
    strSql = "select /*+ Rule*/ * from (" & strSql & ")  res  order by �������,�����,������ϸ,���"
    
    Set ufgMaterialQuery.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(Format(dtpStartDate.value, "yyyy-mm-dd 00:00:00")), _
                                            CDate(Format(dtpEndDate.value, "yyyy-mm-dd 23:59:59")), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text, _
                                            lngCurArchivesId _
                                            )
    
    Call FilterMaterialData
    
End Sub


Private Sub FilterMaterialData()
    Dim strFilter As String
    Dim strState As String
    Dim strIsRequest As String
    
    If ufgMaterialQuery.AdoData Is Nothing Then Exit Sub
    
    strFilter = ""
    strIsRequest = " �Ƿ�����=0"
    
    If cbxStudyType.Text <> "" Then
        strFilter = strFilter & " �������=" & cbxStudyType.ItemData(cbxStudyType.ListIndex)
    End If
    
    If cbxRequestType.Text <> "" Then
        If (cbxRequestType.Text = "����ȡ��" Or cbxRequestType.Text = "������Ƭ") And mcurMaterialType <> amtMaterial Then
        Else
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " ��������='" & cbxRequestType.Text & "'"
            strIsRequest = ""
        End If
    End If
    
    
    If cbxRequestDetail.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " ����ϸĿ='" & cbxRequestDetail.Text & "'"
        strIsRequest = ""
    End If
    
    If Not (chkComplete.value = 0) Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " ִ�й���=6"
    End If
    
    If Not (chkNotEnter.value = 0) Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " ���״̬='δ�鵵'"
    End If
    
    
    strState = ""
    
    '���������Ͳ�Ϊ�����ϣ����飬��Ƭ��ʱ���򲻻�ִ�����¹�������
    If Not (chkWaxStone.value = 0) Then
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " �������='����')"
    End If

    If Not (chkSlices.value = 0) Then
        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " �������='��Ƭ')"
    End If

    If Not (chkTeShu.value = 0) Then
        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " �������='����')"

        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " �������='����')"

        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " �������='��Ⱦ')"
    End If
    
    '�����˵Ĳ������Ͳ��Ǽ�����ʱ����û��ʹ���������ͺ�����ϸĿ�Ĺ�������������Ҫ���˳����Ƿ����롱Ϊ0��������ʾ
    If strIsRequest <> "" And mcurMaterialType <> amtMaterial Then
        strFilter = IIf(strFilter <> "", strFilter & " and " & strIsRequest, strIsRequest)
    End If
        
    ufgMaterialQuery.AdoData.Filter = IIf(strState = "", strFilter, strState)
    
    Call ufgMaterialQuery.RefreshData
End Sub


Private Sub cbxQueryType_KeyPress(KeyAscii As Integer)
'�س�ִ�в�ѯ
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '���ò�ѯ����
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub dtpEndDate_Change()
'�س�ִ�в�ѯ
On Error GoTo errHandle

    '���ò�ѯ����
    Call QueryMaterialData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpStartDate_Change()
'�س�ִ�в�ѯ
On Error GoTo errHandle

    '���ò�ѯ����
    Call QueryMaterialData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtEndPatholNum_KeyPress(KeyAscii As Integer)
'�س�ִ�в�ѯ
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '���ò�ѯ����
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtStartPatholNum_KeyPress(KeyAscii As Integer)
'�س�ִ�в�ѯ
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '���ò�ѯ����
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxRequestDetail_Click()
On Error GoTo errHandle
    If Not cbxRequestDetail.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxRequestType_Click()
On Error GoTo errHandle
    If Not cbxRequestType.Visible Then Exit Sub
    
    Call FilterMaterialData
    
    labRequestDetail.Enabled = True
    cbxRequestDetail.Enabled = True
    Select Case cbxRequestType.Text
        Case "����ȡ��", "��ȡ��", "������Ƭ", "��Ⱦ"
            cbxRequestDetail.ListIndex = 0
            
            labRequestDetail.Enabled = False
            cbxRequestDetail.Enabled = False
        Case "����Ƭ"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("����")
            Call cbxRequestDetail.AddItem("����")
            Call cbxRequestDetail.AddItem("����")
            Call cbxRequestDetail.AddItem("��Ƭ")
            Call cbxRequestDetail.AddItem("��Ⱦ")
            Call cbxRequestDetail.AddItem("��Ƭ")
        Case "����"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("����")
            Call cbxRequestDetail.AddItem("����ҩ")
        Case "����"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("ӫ��")
            Call cbxRequestDetail.AddItem("��ͨ")
        Case Else
            Call ConfigRequestDetail
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxStudyType_Click()
On Error GoTo errHandle
    If Not cbxStudyType.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkComplete_Click()
On Error GoTo errHandle
    If Not chkComplete.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkNotEnter_Click()
On Error GoTo errHandle
    If Not chkNotEnter.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkSlices_Click()
On Error GoTo errHandle
    If Not chkSlices.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkTeShu_Click()
On Error GoTo errHandle
    If Not chkTeShu.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkWaxStone_Click()
On Error GoTo errHandle
    If Not chkWaxStone.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function MaterailEnterArchives(ByVal lngArchivesId As Long) As String
'���������뵵
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim bFind As Boolean
    Dim strLog As String
    Dim strFormId As String
    
    strSql = "select ZL_������_�����뵵([1],[2],[3],[4]) as ����ֵ from dual"
    
    dtServicesTime = zlDatabase.Currentdate
    
    strLog = ""
    For i = 1 To ufgMaterialQuery.GridRows - 1
        If ufgMaterialQuery.GetRowCheck(i) Then
        
            '���жϼ���Ƿ���ɣ�ֻ������ɵļ����ܽ����뵵����
            If Val(ufgMaterialQuery.Text(i, gstrPatholCol_ִ�й���)) = 6 Then
                '�����������Ϊ�����ϣ�����Ҫ�жϲ����Ƿ��Ѿ��뵵�����뵵�Ĳ��ϲ����ٴ��뵵
                If mcurMaterialType = amtMaterial And ufgMaterialQuery.Text(i, gstrPatholCol_���״̬) <> "δ�鵵" Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "�����Ϊ [ " & ufgMaterialQuery.Text(i, gstrPatholCol_�����) & _
                            " ] �Ŀ��Ϊ [ " & ufgMaterialQuery.Text(i, gstrPatholCol_�Ŀ��) & "] ��" & _
                            ufgMaterialQuery.Text(i, gstrPatholCol_������ϸ) & ufgMaterialQuery.Text(i, gstrPatholCol_�������) & "���뵵�������ٴ��뵵��"
                ElseIf mcurMaterialType <> amtMaterial And ufgMaterialQuery.Text(i, gstrPatholCol_���״̬) <> "δ�鵵" Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "�����Ϊ [ " & ufgMaterialQuery.Text(i, gstrPatholCol_�����) & _
                            "] �ļ���ڸõ������Ѿ����ڣ������ٴ��뵵��"
                Else
                
                    strFormId = ufgMaterialQuery.Text(i, gstrPatholCol_��ԴID)
                    strFormId = Mid(strFormId, InStr(strFormId, "-") + 1, 18)
            
                    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                        lngArchivesId, _
                                                        ufgMaterialQuery.Text(i, gstrPatholCol_����ҽ��ID), _
                                                        Val(ufgMaterialQuery.Text(i, gstrPatholCol_������Դ)), _
                                                        Val(strFormId))
    
    
                    If rsData.RecordCount <= 0 Then
                        Call err.Raise(0, "ExecuteArchivesFile", "δ�ɹ���ȡ�뵵����뵵ID,����ʧ�ܡ�")
                        Exit Function
                    End If
                
                    If mcurMaterialType = amtMaterial Then
                        Call ufgMaterialQuery.SyncText(i, gstrPatholCol_���״̬, "�浵��", True)
                    Else
                        Call ufgMaterialQuery.SyncText(i, gstrPatholCol_���״̬, "�Ѵ���", True)
                    End If
                End If
            Else
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "�����Ϊ [ " & ufgMaterialQuery.Text(i, gstrPatholCol_�����) & " ]�ļ����δִ����ɣ����ܽ����뵵������"
            End If
                                                               
        End If
    Next i
    
    MaterailEnterArchives = strLog
End Function


Private Function AllowDelArchivesMatierial(ByVal lngRow As Long) As String
'�жϵ����еĲ����Ƿ������Ƴ�
    AllowDelArchivesMatierial = ""
    
    If ufgArchives.Text(lngRow, gstrPatholCol_����״̬) = ArchivesState_Enter Then
        AllowDelArchivesMatierial = "�����ѹ鵵�����ܴӵ������Ƴ����ϡ�"
        Exit Function
    End If
    
End Function


Private Sub cmdDel_Click()
'ɾ����������
On Error GoTo errHandle
    Dim strInf As String
    
    If Not ufgArchives.IsSelectionRow Then
        Exit Sub
    End If
    
    '�жϸõ����Ƿ������Ƴ�������ʧ���ѽ��ĵĲ��ϲ��ܽ����Ƴ���������
    strInf = AllowDelArchivesMatierial(ufgArchives.SelectionRow)
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�Ƴ��ĵ������ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "ȷ��Ҫ�ӵ������Ƴ���ѡ��Ĳ�����", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    strInf = Execute_ClearArchivesMaterial
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
    End If
    
    
    Call RefreshStateInf(False, True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdEnterArchives_Click()
'�����뵵
On Error GoTo errHandle
    Dim strLog As String
    
    If mlngCurArchivesId <= 0 Then
        Call MsgBoxD(Me, "��ѡ�������ĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgMaterialQuery.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����뵵�����ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strLog = MaterailEnterArchives(Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID)))
    
    If strLog <> "" Then
        Call MsgBoxD(Me, strLog, vbOKOnly, Me.Caption)
    Else
        Call MsgBoxD(Me, "������뵵������", vbOKOnly, Me.Caption)
    End If
    
    Call RefreshStateInf(False, True)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintArchives(ByVal lngArchivesId As Long, ByVal strReportName As String, Optional ByVal blnIsPrint As Boolean = True)
'��ӡ��������
    Dim i As Long
    Dim j As Long
    Dim strValue(7) As String
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0": strValue(6) = "0": strValue(7) = "0"
    
    If mcurMaterialType = amdReport Then
        For i = 1 To ufgArchivesDetail.GridRows - 1
            If ufgArchivesDetail.GetRowCheck(i) Then
                If blnIsPrint Then
                    Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
                        "����ID=" & lngArchivesId, "�鵵ID=" & ufgArchivesDetail.KeyValue(i), 2)
                Else
                    '�����Ԥ������ֻԤ����һ��ѡ�е�����
                    Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
                        "����ID=" & lngArchivesId, "�鵵ID=" & ufgArchivesDetail.KeyValue(i), 1)
                        
                    Exit Sub
                End If
            End If
        Next i

    Else
        For i = 1 To ufgArchivesDetail.GridRows - 1
            If ufgArchivesDetail.GetRowCheck(i) Then
                If zlCommFun.ActualLen(strValue(j)) > 3000 Then
                    j = j + 1
                    strValue(j) = ""
                End If
    
                If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
                strValue(j) = strValue(j) & ufgArchivesDetail.KeyValue(i)
            End If
        Next i
        
        Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
            "����ID=" & lngArchivesId, "�鵵ID1=" & strValue(0), "�鵵ID2=" & strValue(1), "�鵵ID3=" & strValue(2), "�鵵ID4=" & strValue(3), "�鵵ID5=" & strValue(4), "�鵵ID6=" & strValue(5), "�鵵ID7=" & strValue(6), "�鵵ID8=" & strValue(7), _
            IIf(blnIsPrint, 2, 1)) '1��Ԥ����2����ӡ
    End If
    

End Sub


Private Sub cmdFilter_Click()
On Error GoTo errHandle
    Dim strFilter As String
    
    If ufgArchivesDetail.AdoData Is Nothing Then
        Call cmdRead_Click
    End If
    
    Call frmPatholArchivesLocate.ShowFilterWindow(Me)
    
    If Not frmPatholArchivesLocate.blnOk Then Exit Sub
        

    strFilter = " ����ʱ��>='" & Format(frmPatholArchivesLocate.dtpStart.value, "yyyy-mm-dd 00:00:00") & "' and ����ʱ�� <= '" & Format(frmPatholArchivesLocate.dtpEnd.value, "yyyy-mm-dd 23:59:59") & "'"


    If frmPatholArchivesLocate.txtName.Text <> "" Then
        strFilter = " ���� like '" & frmPatholArchivesLocate.txtName.Text & "*'"
    End If
    
    
    If frmPatholArchivesLocate.txtPatholNum.Text <> "" Then
        strFilter = " �����='" & frmPatholArchivesLocate.txtPatholNum.Text & "' or �����='" & UCase(frmPatholArchivesLocate.txtPatholNum.Text) & "'"
    End If
        

    
    ufgArchivesDetail.AdoData.Filter = strFilter
    
    Call ufgArchivesDetail.RefreshData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdPreview_Click()
'��ӡ��������
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������) = "" Then
        Call MsgBoxD(Me, "��������δ���ö�Ӧ�������ڵ����������������ö�Ӧ�ı������ơ�", vbOKOnly, Me.Caption)
'        If MsgBoxD(Me, "��������δ���ö�Ӧ�������ڵ����������������ö�Ӧ�ı������ơ��Ƿ��������ã�", vbYesNo, Me.Caption) = vbNo Then
'            Exit Sub
'        Else
'            Call mnu_ArchivesClassCfg_Click
'        End If
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����ҪԤ���ĵ������ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call PrintArchives(mlngCurArchivesId, ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������), False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdPrint_Click()
'��ӡ��������
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������) = "" Then
        Call MsgBoxD(Me, "��������δ���ö�Ӧ�������ڵ����������������ö�Ӧ�ı������ơ�", vbOKOnly, Me.Caption)
'        If MsgBoxD(Me, "��������δ���ö�Ӧ�������ڵ����������������ö�Ӧ�ı������ơ��Ƿ��������ã�", vbYesNo, Me.Caption) = vbNo Then
'            Exit Sub
'        Else
'            Call mnu_ArchivesClassCfg_Click
'
'        End If
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ĵ������ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call PrintArchives(mlngCurArchivesId, ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������), True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdQuery_Click()
'��ѯ�鵵����
On Error GoTo errHandle
    Call QueryMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdRead_Click()
On Error GoTo errHandle

    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    Call LoadArchivesDetail(mlngCurArchivesId)
    
    If ufgArchivesDetail.AdoData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "�õ����в�������ϸ���ݡ�", vbOKOnly, Me.Caption)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadParameterConfig()
'������ز�������
    mlngDefaultQueryDays = zlDatabase.GetPara("����Ĭ�ϲ�ѯ����", glngSys, G_LNG_PATHOLARCHIVES_NUM, "30")
    mstrLabelReportName = zlDatabase.GetPara("������ǩ��������", glngSys, G_LNG_PATHOLARCHIVES_NUM, "")
End Sub


Private Sub ConfigPopedomFace()
'����Ȩ�����ý��棬������߱�Ȩ��ʱ�������ض�Ӧ���ܰ�ť
    Dim i As Long
    
    mnu_ParameterConfig.Visible = CheckPopedom(mstrPrivs, "��������")
    mnu_ArchivesClassCfg.Visible = CheckPopedom(mstrPrivs, "��������")
    
    mnu_CancelArchives.Visible = CheckPopedom(mstrPrivs, "�����鵵")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_CancelArchives") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "�����鵵")
        End If
    Next i
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
'    #If DebugState = True Then
'        Call InitDebugObject(1295, Me, "zlhis", "HIS")
'    #End If
    mblnIsFormLoaded = False
    
    Call RestoreWinState(Me, App.ProductName)
    
'    Call InitCommandBars
    
    Call InitFace
    Call InitMenuIcoConfig
    
    Call InitArchivesFileList
    
    Call LoadParameterConfig
    
    Call ConfigStudyType
    Call ConfigRequestType
    Call ConfigRequestDetail
    
    Call SwitchArchivesFace(amtTable)
    
    curDate = zlDatabase.Currentdate
    
    cbxQueryType.ListIndex = 0
    dtpStartDate.value = curDate
    dtpEndDate.value = curDate
    mlngCurArchivesId = -1
    
    mstrPrivs = gstrPrivs
    
    Call ConfigPopedomFace
    
    Set zlReport = New zl9Report.clsReport
    
    
    Call QueryArchivesData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
    Call RefreshStateInf(True, True)
    mblnIsFormLoaded = True
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'Private Sub InitCommandBars()
'    '���ܴ���������
'    Dim cbrControl As CommandBarControl
'    Dim cbrPopControl As CommandBarControl
'    Dim cbrMenuBar As CommandBarPopup
'    Dim cbrToolBar As CommandBar
'    Dim cbrCustom As CommandBarControlCustom
'    Dim str3DFuncs() As String
'
'    Dim rsCollection As ADODB.Recordset
'    Dim rsViewShare As ADODB.Recordset
'    Dim rsShareCount As ADODB.Recordset
'    Dim rsTemp As ADODB.Recordset
'
'    Dim i As Integer
'    Dim i3DFunc As Integer
'
'    '-----------------------------------------------------
'    CommandBarsGlobalSettings.App = App
'    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
'    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
'
'    Me.cbrQuery.VisualTheme = xtpThemeOffice2003
'
'    Set Me.cbrQuery.Icons = zlCommFun.GetPubIcons
'    With Me.cbrQuery.Options
'        .ShowExpandButtonAlways = False
'        .ToolBarAccelTips = True
'        .AlwaysShowFullMenus = False
'        .IconsWithShadow = True '����VisualTheme����Ч
'        .UseDisabledIcons = True
'        .LargeIcons = True
'        .SetIconSize True, 24, 24
'    End With
'    Me.cbrQuery.EnableCustomization False
'    Me.cbrQuery.ActiveMenuBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'
'
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 1, "��ѯʱ��")
'        cbrCustom.Handle = cbxQueryType.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 2, "��ʼʱ��")
'        cbrCustom.Handle = dtpStartDate.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 3, "��")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 4, "����ʱ��")
'        cbrCustom.Handle = dtpEndDate.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 5, "����ţ�")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 6, "��ʼ�����")
'        cbrCustom.Handle = txtStartPatholNum.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 7, "��")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 8, "���������")
'        cbrCustom.Handle = txtEndPatholNum.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 9, "��ѯ��ť")
'        cbrCustom.Handle = cmdQuery.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
''    Set cbrToolBar = Me.cbrQuery.Add("������", xtpBarTop)
''    cbrToolBar.ShowTextBelowIcons = True
'
''    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
''    With cbrToolBar.Controls
''        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "�� ѯ")
''            cbrControl.IconId = 814
''            cbrControl.ToolTipText = "�� ѯ"
'
''        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
''        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
''        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����"): cbrControl.IconId = 744
''
''        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
''        cbrControl.BeginGroup = True
''
''    End With
'
'End Sub


Private Sub ConfigStudyType()
'���ü������
    Call cbxStudyType.Clear

    Call cbxStudyType.AddItem("")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = -1
        
    Call cbxStudyType.AddItem("����")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 0
    
    Call cbxStudyType.AddItem("����")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 1
    
    Call cbxStudyType.AddItem("ϸ��")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 2
    
    Call cbxStudyType.AddItem("����")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 3
    
    Call cbxStudyType.AddItem("ʬ��")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 4
    
    Call cbxStudyType.AddItem("����ʯ��")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 5
    
    cbxStudyType.ListIndex = 0
End Sub


Private Sub ConfigRequestType()
'������������
    Call cbxRequestType.Clear
    
    Call cbxRequestType.AddItem("")
    
    Call cbxRequestType.AddItem("����ȡ��")
    Call cbxRequestType.AddItem("��ȡ��")
    Call cbxRequestType.AddItem("������Ƭ")
    Call cbxRequestType.AddItem("����Ƭ")
    Call cbxRequestType.AddItem("����")
    Call cbxRequestType.AddItem("��Ⱦ")
    Call cbxRequestType.AddItem("����")
    
    
    cbxRequestType.ListIndex = 0
End Sub


Private Sub ConfigRequestDetail()
'��������ϸĿ
    Call cbxRequestDetail.Clear
    
    Call cbxRequestDetail.AddItem("")
    
    Call cbxRequestDetail.AddItem("����")
    Call cbxRequestDetail.AddItem("����ҩ")
    
    Call cbxRequestDetail.AddItem("ӫ��")
    Call cbxRequestDetail.AddItem("��ͨ")
    
    Call cbxRequestDetail.AddItem("����")
    Call cbxRequestDetail.AddItem("����")
    Call cbxRequestDetail.AddItem("����")
    Call cbxRequestDetail.AddItem("��Ƭ")
    Call cbxRequestDetail.AddItem("��Ⱦ")
    Call cbxRequestDetail.AddItem("��Ƭ")
    
    cbxRequestDetail.ListIndex = 0
End Sub


Private Sub SwitchArchivesFace(ByVal amtMaterialType As TArchivesMaterialType)
'���ݲ��������л��������Ͻ���
'    If mcurMaterialType = amtMaterialType Then Exit Sub
    
    mcurMaterialType = amtMaterialType
    
    Call InitArchivesQueryList(amtMaterialType)
    Call InitArchivesDetailList(amtMaterialType)
    
    txtNumberInf.Visible = IIf(amtMaterialType = amtMaterial, True, False) And tabFilter.Selected.Index = 1
    
    labRequestDetail.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    cbxRequestDetail.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    lineSplit2.Visible = IIf(amtMaterialType = amtMaterial, True, False)
'    chkNotEnter.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkWaxStone.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkSlices.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkTeShu.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    
    
    cbxRequestType.ListIndex = 0
    cbxRequestDetail.ListIndex = 0
    chkNotEnter.value = 1 'IIf(mcurMaterialType = amtMaterial, 1, 0)
    chkComplete.value = 1
    chkWaxStone.value = 0
    chkSlices.value = 0
    chkTeShu.value = 0
    
    Call Picture2_Resize
End Sub


Private Sub InitArchivesFileList()
'��ʼ�������б�
    Dim strTemp As String
    

    
    ufgArchives.IsKeepRows = False
    ufgArchives.IsCopyMode = True
    
    strTemp = zlDatabase.GetPara("�����б�����", glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    If strTemp = "" Then
        ufgArchives.ColNames = gstrArchivesManageCols
    Else
        ufgArchives.ColNames = strTemp
    End If
        '��������
    ufgArchives.GridRows = glngStandardRowCount
    '�����и�
    ufgArchives.RowHeightMin = glngStandardRowHeight
    ufgArchives.DefaultColNames = gstrArchivesManageCols
    ufgArchives.ColConvertFormat = gstrArchivesManageConvertFormat
End Sub


Private Sub InitArchivesQueryList(ByVal amtMaterialType As TArchivesMaterialType)
'��ʼ����������б�
'lngMaterialType:��������  0-���ֲ��ϣ�1-������

    Dim strTemp As String

    
    strTemp = zlDatabase.GetPara(IIf(mcurMaterialType = amtMaterial, "�������ϲ�ѯ�б�����", "����ֽ�ʲ�ѯ�б�����"), glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    ufgMaterialQuery.IsKeepRows = False
    If amtMaterialType <> amtMaterial Then
        ufgMaterialQuery.IsCopyMode = True
        ufgMaterialQuery.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesWordCols)
        ufgMaterialQuery.DefaultColNames = gstrArchivesWordCols
        ufgMaterialQuery.ColConvertFormat = gstrArchivesWordConvertFormat
    Else
        ufgMaterialQuery.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesMaterialCols)
        ufgMaterialQuery.DefaultColNames = gstrArchivesMaterialCols
        ufgMaterialQuery.ColConvertFormat = gstrArchivesMaterialConvertFormat
    End If
        
    '��������
    ufgMaterialQuery.GridRows = glngStandardRowCount
    '�����и�
    ufgMaterialQuery.RowHeightMin = glngStandardRowHeight
    Set ufgMaterialQuery.AdoData = Nothing
    Call ufgMaterialQuery.RefreshData
    
End Sub



Private Sub InitArchivesDetailList(ByVal amtMaterialType As TArchivesMaterialType)
'��ʼ����������б�
'lngMaterialType:��������  0-���ֲ��ϣ�1-������
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara(IIf(mcurMaterialType = amtMaterial, "����������ϸ�б�����", "����ֽ����ϸ�б�����"), glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    ufgArchivesDetail.IsKeepRows = False
    If amtMaterialType <> amtMaterial Then
        ufgArchivesDetail.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesWordCols)
        ufgArchivesDetail.DefaultColNames = gstrArchivesWordCols
        ufgArchivesDetail.ColConvertFormat = gstrArchivesWordConvertFormat
    Else
        ufgArchivesDetail.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesMaterialDetailCols)
        ufgArchivesDetail.DefaultColNames = gstrArchivesMaterialDetailCols
        ufgArchivesDetail.ColConvertFormat = gstrArchivesMaterialConvertFormat
    End If
        '��������
    ufgArchivesDetail.GridRows = glngStandardRowCount
    '�����и�
    ufgArchivesDetail.RowHeightMin = glngStandardRowHeight
    Set ufgArchivesDetail.AdoData = Nothing
    Call ufgArchivesDetail.RefreshData
End Sub


Private Sub QueryArchivesData(ByVal dtStartDate As Date, ByVal dtEndDate As Date, Optional ByVal lngArchivesClassId As Long, _
    Optional ByVal strArchivsName As String, Optional ByVal strArchivesCode As String)
'��ѯָ��ʱ�䷶Χ�ڵ����ݵ�����
    Dim strSql As String
    
    
    mblnMoved = MovedByDate(dtStartDate)
    
    strSql = "select a.ID, a.��������, a.�������, a.��鷶Χ, " & _
                " a.��ʼ����, a.��������, a.����˵��, a.����״̬, a.������, a.��������, b.�������� as ��������, B.��������,B.��������," & _
                " a.��������, a.�������, a.��������, a.��ϸ��ַ,a.�鵵ʱ�� " & _
                " from ��������Ϣ a, ���������� b " & _
                " where a.����ID=b.id  and a.�������� between [1] and [2] " & _
                IIf(lngArchivesClassId <= 0, "", " and a.����ID=[3]") & _
                IIf(strArchivsName = "", "", " and upper(a.��������)=upper([4])") & _
                IIf(strArchivesCode = "", "", " and upper(a.�������)=upper([5])") & _
                " order by a.��������,a.�������� "
                   
    Set ufgArchives.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtStartDate, "yyyy-mm-dd 00:00:00")), _
                            CDate(Format(dtEndDate, "yyyy-mm-dd 23:59:59")), lngArchivesClassId, strArchivsName, strArchivesCode)
    
    Call ufgArchives.RefreshData
    
    Call ufgArchives.LocateRow(1)
    
    '��ȡ����˵����Ϣ
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
End Sub


Private Sub InitFace()
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
        

        .InsertItem 0, "�����뵵", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "�����뵵"
        
        
        .InsertItem 1, "������ϸ", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "������ϸ"
        
        .Item(0).Selected = True
    End With
    
    framEnterArchives.Visible = True
End Sub



Private Sub AdjustLayOut()
    
    Picture1.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    Picture1.Height = Me.ScaleHeight - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
    
    Call ucSplitter1.RePaint
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustLayOut
err.Clear
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Set zlReport = Nothing
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

Private Sub mnu_ArchivesClassCfg_Click()
On Error GoTo errHandle
    If Not CheckPopedom(mstrPrivs, "��������") Then
        Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim frmArchivesClass As New frmPatholArchivesClass
    On Error GoTo errFree
        Call frmArchivesClass.Show(1, Me)
errFree:
        Call Unload(frmArchivesClass)
        Set frmArchivesClass = Nothing
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

Private Sub mnu_CancelArchives_Click()
'�����鵵
On Error GoTo errHandle
    Call Execute_CancelEnterArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_DelArchives_Click()
'ɾ������
On Error GoTo errHandle
    Call Execute_DelArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_EnterArchives_Click()
'�����鵵
On Error GoTo errHandle
    Call Execute_EnterArchives
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



Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��ӡԤ��, 0��������ѡ��Ի���1Ԥ����2��ӡ��3����Excel
    '������    �����ʽ
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgArchives.DataGrid
    
    objPrint.Title = "�������嵥"

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



Private Sub mnu_Font_Click()
'����
On Error GoTo errHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgArchives.DataGrid.Font.Bold
    diaFont.FontName = ufgArchives.DataGrid.Font.Name
    diaFont.FontSize = ufgArchives.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgArchives.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgArchives.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '�����б�
    ufgArchives.DataGrid.Font.Bold = diaFont.FontBold
    ufgArchives.DataGrid.Font.Name = diaFont.FontName
    ufgArchives.DataGrid.Font.Size = diaFont.FontSize
    ufgArchives.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgArchives.DataGrid.Font.Underline = diaFont.FontUnderline
    
    
    Call ufgArchives.DataGrid.Refresh
    
    ufgArchives.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgArchives.DataGrid.AutoSize(0, ufgArchives.DataGrid.Rows - 1)
    
    ufgArchives.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgArchives.DataGrid.AutoSize(0, ufgArchives.DataGrid.Rows - 1)
    
    
    '��ѯ�б�
    ufgMaterialQuery.DataGrid.Font.Bold = diaFont.FontBold
    ufgMaterialQuery.DataGrid.Font.Name = diaFont.FontName
    ufgMaterialQuery.DataGrid.Font.Size = diaFont.FontSize
    ufgMaterialQuery.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgMaterialQuery.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgMaterialQuery.DataGrid.Refresh
    
    ufgMaterialQuery.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgMaterialQuery.DataGrid.AutoSize(0, ufgMaterialQuery.DataGrid.Rows - 1)
    
    ufgMaterialQuery.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgMaterialQuery.DataGrid.AutoSize(0, ufgMaterialQuery.DataGrid.Rows - 1)
    
    
    '��ϸ�б�
    ufgArchivesDetail.DataGrid.Font.Bold = diaFont.FontBold
    ufgArchivesDetail.DataGrid.Font.Name = diaFont.FontName
    ufgArchivesDetail.DataGrid.Font.Size = diaFont.FontSize
    ufgArchivesDetail.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgArchivesDetail.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgArchivesDetail.DataGrid.Refresh
    
    ufgArchivesDetail.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgArchivesDetail.DataGrid.AutoSize(0, ufgArchivesDetail.DataGrid.Rows - 1)
    
    ufgArchivesDetail.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgArchivesDetail.DataGrid.AutoSize(0, ufgArchivesDetail.DataGrid.Rows - 1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Help_Click()
'����
On Error GoTo errHandle
    Call Execute_Help
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_LabPreview_Click()
'��ǩԤ��
On Error GoTo errHandle
    Call Execute_PrintArchivesLabel(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_LabPrint_Click()
'��ǩ��ӡ
On Error GoTo errHandle
    Call Execute_PrintArchivesLabel(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub mnu_ListPreview_Click()
'Ԥ�������б�
On Error GoTo errHandle
    Call MenuPrint(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ListPrint_Click()
'��ӡ�����б�
On Error GoTo errHandle
    Call MenuPrint(1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MainPage_Click()
'������ҳ
On Error GoTo errHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_NewArchives_Click()
'��������
On Error GoTo errHandle
    Call Execute_NewArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ParameterConfig_Click()
'��������
On Error GoTo errHandle
    If Not CheckPopedom(mstrPrivs, "��������") Then
        Call MsgBoxD(Me, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Execute_ParameterConfig
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintConfig_Click()
'��ӡ����
On Error GoTo errHandle
    Call zlPrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_QueryArchives_Click()
'��ѯ����
On Error GoTo errHandle
    Call Execute_QueryArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Return_Click()
'���ͷ���
On Error GoTo errHandle
    Call zlMailTo(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StandardBut_Click()
On Error GoTo errHandle
    Dim intCount As Long
    Me.mnu_StandardBut.Checked = Not Me.mnu_StandardBut.Checked
    Me.tbrTools.Visible = Me.mnu_StandardBut.Checked
    
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
    Picture3.Left = 120
    Picture3.Top = 120
    Picture3.Width = Picture1.ScaleWidth - 120
    Picture3.Height = Picture1.ScaleHeight - 120
    
    Call ucSplitter2.RePaint
err.Clear
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    tabFilter.Left = 0
    tabFilter.Top = 0
    tabFilter.Width = Picture2.ScaleWidth
    
    framEnterArchives.Left = 0
    framEnterArchives.Top = tabFilter.Height
    framEnterArchives.Width = Picture2.ScaleWidth
    framEnterArchives.Height = Picture2.ScaleHeight - tabFilter.Height
    
    framQuery.Width = framEnterArchives.Width
    
    
    
    
    ufgMaterialQuery.Left = 120
    ufgMaterialQuery.Top = cbxStudyType.Top + cbxStudyType.Height + 120
    ufgMaterialQuery.Width = framEnterArchives.Width - 240
    ufgMaterialQuery.Height = framEnterArchives.Height - cbxStudyType.Top - chkComplete.Height - 840   '- cmdEnterArchives.Height
    
    
'    If cbxRequestDetail.Visible Then
'        chkComplete.Left = cbxRequestDetail.Left + cbxRequestDetail.Width + 120
'        chkComplete.Top = cbxRequestDetail.Top + 50
'    Else
'        chkComplete.Left = cbxRequestType.Left + cbxRequestType.Width + 120
'        chkComplete.Top = cbxRequestType.Top + 50
'    End If

    cmdEnterArchives.Left = 120 'framQuery.Width - cmdEnterArchives.Width - 120
    cmdEnterArchives.Top = ufgMaterialQuery.Top + ufgMaterialQuery.Height + 120

    If chkTeShu.Visible Then
        chkComplete.Left = ufgMaterialQuery.Width - chkComplete.Width * 5 - 600 '120
    Else
        chkComplete.Left = ufgMaterialQuery.Width - chkComplete.Width * 2 - 240
    End If
    
    chkComplete.Top = cmdEnterArchives.Top + 50
    
    chkNotEnter.Left = chkComplete.Left + chkComplete.Width + 120
    chkNotEnter.Top = chkComplete.Top
    
    lineSplit2.X1 = chkNotEnter.Left + chkNotEnter.Width + 120
    lineSplit2.X2 = lineSplit2.X1
    lineSplit2.Y1 = chkNotEnter.Top
    lineSplit2.Y2 = lineSplit2.Y1 + chkNotEnter.Height
    
    chkWaxStone.Left = lineSplit2.X1 + 120
    chkWaxStone.Top = chkComplete.Top
    
    chkSlices.Left = chkWaxStone.Left + chkWaxStone.Width + 120
    chkSlices.Top = chkComplete.Top
    
    chkTeShu.Left = chkSlices.Left + chkSlices.Width + 120
    chkTeShu.Top = chkComplete.Top
    
    
    
    '================================================================================
    
    framArchivesDetail.Left = 0
    framArchivesDetail.Top = tabFilter.Height
    framArchivesDetail.Width = Picture2.ScaleWidth
    framArchivesDetail.Height = Picture2.ScaleHeight - tabFilter.Height
    
    ufgArchivesDetail.Left = 120
    ufgArchivesDetail.Top = 240
    ufgArchivesDetail.Width = framArchivesDetail.Width - 240
    ufgArchivesDetail.Height = framArchivesDetail.Height - cmdDel.Height - 480
        
    cmdDel.Left = framArchivesDetail.Width - cmdDel.Width - 120
    cmdDel.Top = ufgArchivesDetail.Top + ufgArchivesDetail.Height + 120
    
    cmdPrint.Left = cmdDel.Left - cmdPrint.Width - 120
    cmdPrint.Top = cmdDel.Top
    
    cmdPreview.Left = cmdPrint.Left - cmdPreview.Width - 120
    cmdPreview.Top = cmdDel.Top
    
    cmdRead.Left = 120
    cmdRead.Top = cmdDel.Top
    
    cmdFilter.Left = cmdRead.Left + cmdRead.Width + 120
    cmdFilter.Top = cmdDel.Top
err.Clear
End Sub



Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    
    framEnterArchives.Visible = IIf(Item.Index = 0, True, False)
    framArchivesDetail.Visible = IIf(Item.Index = 0, False, True)
    txtNumberInf.Visible = IIf(Item.Index = 0, False, True) And mcurMaterialType = amtMaterial
    
'    If Item.Index = 1 Then
'        If Not ufgArchives.IsSelectRow Then Exit Sub
'        If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectRowIndex, gstrArchivesManage_ID))
'
'        Call LoadArchivesDetail(mlngCurArchivesId)
'    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function AllowUpdateArchivesFile(ByVal lngDelRow As Long) As String
'�ж��Ƿ�������µ���
    AllowUpdateArchivesFile = ""
    
    If mblnMoved Then
        AllowUpdateArchivesFile = "�����ѱ�ת�ƣ����ܽ��и��¡�"
        Exit Function
    End If
    
    If ufgArchives.Text(lngDelRow, gstrPatholCol_����״̬) <> "δ�鵵" Then
        AllowUpdateArchivesFile = "�����ѹ鵵�����ܽ��и��¡�"
        Exit Function
    End If
End Function


Private Function AllowDelArchivesFile(ByVal lngDelRow As Long) As String
'�ж��Ƿ�����ɾ������
    AllowDelArchivesFile = ""
    
    If mblnMoved Then
        AllowDelArchivesFile = "�����ѱ�ת�ƣ����ܽ���ɾ����"
        Exit Function
    End If
    
    If ufgArchives.Text(lngDelRow, gstrPatholCol_����״̬) <> "δ�鵵" Then
        AllowDelArchivesFile = "�����ѹ鵵�����ܽ���ɾ����"
        Exit Function
    End If
    
'    If ufgStudy.ShowDataRows > 0 Then
'        AllowDelArchivesFile = "�����а���������ݣ����ܽ���ɾ����"
'        Exit Function
'    End If
End Function


Private Sub DelArchivesFileData(lngArchivesId As Long)
'ɾ��������¼
    Dim strSql As String
    
    strSql = "Zl_������_ɾ���ļ�����(" & lngArchivesId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub


Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errHandle
    
    Select Case UCase(Button.Key)
        Case UCase("tbn_LabView")   'Ԥ��������ǩ
            Call Execute_PrintArchivesLabel(False)
            
        Case UCase("tbn_LabPrint")  '��ӡ������ǩ
            Call Execute_PrintArchivesLabel(True)
            
        Case UCase("tbn_NewArchives")   '��������
            Call Execute_NewArchives
    
        Case UCase("tbn_DelArchives")   'ɾ������
            Call Execute_DelArchives
            
        Case UCase("tbn_UpdateArchives")    '���µ���
            Call Execute_UpdateArchives
                
        Case UCase("tbn_QueryArchives")     '��ѯ����
            Call Execute_QueryArchives
            
        Case UCase("tbn_EnterArchives")     '�����鵵
            Call Execute_EnterArchives
            
        Case UCase("tbn_CancelArchives")    '�����鵵
            Call Execute_CancelEnterArchives
            
        Case UCase("tbn_Help")  '����
            Call Execute_Help
            
        Case UCase("tbn_Exit")  '�Ƴ���������ģ��
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Execute_Exit()
'�˳�
    Call Unload(Me)
End Sub


Private Sub Execute_Help()
'����
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Execute_QueryArchives()
'��ѯ����
On Error GoTo errHandle
    Dim strSql As String
    
    Call frmPatholArchivesQuery.ShowArchivesQueryWindow(mlngDefaultQueryDays, Me)
    
    If frmPatholArchivesQuery.mblnIsOk Then
        Call QueryArchivesData(frmPatholArchivesQuery.dtStartDate, frmPatholArchivesQuery.dtEndDate, _
            frmPatholArchivesQuery.lngArchivesClassId, frmPatholArchivesQuery.strArchivesName, frmPatholArchivesQuery.strArchivesCode)
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Execute_NewArchives()
'��������
    If Not frmPatholArchivesFileNew.ShowAddArchivesFileWindow(ufgArchives, Me) Then Exit Sub
    
    '��ȡ����������ʾ��Ϣ
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_DelArchives()
'ɾ������
    Dim strInf As String
    
    '��Ҫ�жϵ����Ƿ��Ѿ���棬�ҵ����в��������
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    strInf = AllowDelArchivesFile(ufgArchives.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ��ѡ��ĵ�����¼��", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    Call DelArchivesFileData(mlngCurArchivesId)
    Call ufgArchives.DelRow(ufgArchives.SelectionRow, False, True)
    
    '��ȡ����������ʾ��Ϣ
    If ufgArchives.SelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    Else
        '...
    End If
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_UpdateArchives()
'���µ���
    Dim strInf As String
    
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���µĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = AllowUpdateArchivesFile(ufgArchives.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call frmPatholArchivesFileNew.ShowUpdateArchivesFileWindow(ufgArchives, Me)
    
    '��ȡ����������ʾ��Ϣ
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
End Sub


Private Function ShowPlaceSureWindow(ByVal lngArchivesIndex As Long, ByRef strRoom As String, _
                                ByRef strBox As String, ByRef strDrawer As String) As Boolean
    Dim frmPlaceDialog As frmPatholArchivesPlaceDialog
    
    strRoom = ""
    strBox = ""
    strDrawer = ""
    
    On Error GoTo errFree:
        Set frmPlaceDialog = New frmPatholArchivesPlaceDialog
        
        Call frmPlaceDialog.ShowPlaceDialog(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������), _
                                        ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_�������), _
                                        ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������), _
                                        Me)
        If frmPlaceDialog.IsOk Then
            strRoom = frmPlaceDialog.Room
            strBox = frmPlaceDialog.Box
            strDrawer = frmPlaceDialog.Drawer
        End If
        
        ShowPlaceSureWindow = frmPlaceDialog.IsOk
errFree:
    Call Unload(frmPlaceDialog)
    Set frmPlaceDialog = Nothing
End Function

Private Sub Execute_EnterArchives()
'ִ�е����鵵����
    Dim strRoom As String
    Dim strBox As String
    Dim strDrawer As String
    Dim curDate As Date
    
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�鵵�ĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_����״̬) = ArchivesState_Enter Then
        Call MsgBoxD(Me, "�����ѹ鵵�����ܽ��й鵵����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ShowPlaceSureWindow(ufgArchives.SelectionRow, strRoom, strBox, strDrawer) Then Exit Sub
    
    If strDrawer = "" And strBox = "" And strRoom = "" Then
        Call MsgBoxD(Me, "δѡ�񵵰����λ�ã����ܽ��й鵵��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
        
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    '���µ������λ��
    Call zlDatabase.ExecuteProcedure("ZL_������_λ�ø���(" & mlngCurArchivesId & _
                                        ",'" & strRoom & "','" & strBox & "','" & strDrawer & "')", Me.Caption)
    '���µ���״̬
    curDate = zlDatabase.Currentdate
    
    Call zlDatabase.ExecuteProcedure("Zl_������_�ļ������鵵(" & mlngCurArchivesId & ",1," & To_Date(Format(curDate, "yyyy-mm-dd")) & ")", Me.Caption)
    
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_����״̬) = "�ѹ鵵"
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������) = strRoom
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_�������) = strBox
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������) = strDrawer
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_�鵵ʱ��) = Format(curDate, "yyyy-mm-dd")

    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(True)
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_ParameterConfig()
'��������
    Dim frmParameter As frmPatholArchivesParameter
    
    Set frmParameter = New frmPatholArchivesParameter
On Error GoTo errFree
    Call frmParameter.ShowParameterWindow(mlngDefaultQueryDays, mstrLabelReportName, Me)
    
    mlngDefaultQueryDays = frmParameter.lngDefaultQueryDays
    mstrLabelReportName = frmParameter.strLabelReportName
errFree:
    Call Unload(frmParameter)
    Set frmParameter = Nothing
    
End Sub

Private Sub Execute_PrintArchivesLabel(ByVal blnIsAtOncePrint As Boolean)
'Ԥ����ӡ������ǩ
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ��ӡ�ĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Trim(mstrLabelReportName) = "" Then
        Call MsgBoxD(Me, "��δ���ñ�ǩ��Ӧ�ı������ƣ��뵽���������á��н������á�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
        
    Call zlReport.ReportOpen(gcnOracle, 100, mstrLabelReportName, Me, "����ID=" & mlngCurArchivesId, IIf(blnIsAtOncePrint, 2, 1)) '1��Ԥ����2����ӡ
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Execute_CancelEnterArchives()
'ִ�е��������鵵����

    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ�����鵵�ĵ�����¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "�����ѱ�ת�ƣ�����ִ�иò�����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_����״̬) = ArchivesState_NoEnter Then
        Call MsgBoxD(Me, "����δ�鵵�����ܽ��г�������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫ�Ըõ������г����鵵�Ĳ����𣿳����鵵�󣬵��������Ϣ�������޸ġ�", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    Call zlDatabase.ExecuteProcedure("Zl_������_�ļ������鵵(" & mlngCurArchivesId & ",0,null)", Me.Caption)
    
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_����״̬) = "δ�鵵"
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_�鵵ʱ��) = ""
    
    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(False)
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub LoadArchivesDetail(ByVal lngArchivesId As Long)
    Dim strSql As String
    
    If lngArchivesId <= 0 Then Exit Sub
    
    If mcurMaterialType <> amtMaterial Then
        strSql = "select /*+ Rule*/ * from (" & _
                " select a.id as ��ԴID, a.����ҽ��ID, 4 as ������Դ, b.�����,c.����,c.�Ա�,c.����,c.ҽ������ as �����Ŀ, b.�������, " & _
                " decode(a.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') as ���״̬, decode(a.����״̬, 0, 'δ���', 1, '���ֽ��', '�ѽ��') as ����״̬,a.����״̬ as ����,b.����ʱ��,null as ִ�й��� " & _
                " from ����鵵��Ϣ a,  ��������Ϣ b, ����ҽ����¼ c " & _
                " Where a.������Դ = 4 And a.����ҽ��id = b.����ҽ��id And b.ҽ��ID = c.ID and ����ID=[1]" & _
                " )order by ����"
    Else
        strSql = "select /*+ Rule*/ * from (" & _
                " select a.id as ��ԴID, a.����ҽ��ID, 1 as ������Դ, c.�����,d.����,d.�Ա�,d.����,d.ҽ������ as �����Ŀ, c.�������, b.���, b.�걾����, b.ȡ��λ��, '����' as �������, " & _
                " case when b.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, b.������ as ����, " & _
                " decode(a.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') as ���״̬, decode(a.����״̬, 0, 'δ���', 1, '���ֽ��', '�ѽ��') as ����״̬,a.����״̬ as ����,c.����ʱ��, null as ִ�й��� " & _
                " from ����鵵��Ϣ a, ����ȡ����Ϣ b, ��������Ϣ c, ����ҽ����¼ d " & _
                " Where a.������Դ = 1 And a.�Ŀ�ID = b.�Ŀ�ID And b.����ҽ��id = c.����ҽ��id And c.ҽ��ID = d.ID and ����ID=[1] " & _
            " Union All " & _
                " select a.id as ��ԴID, a.����ҽ��ID, 2 as ������Դ, d.�����,e.����,e.�Ա�,e.����,e.ҽ������ as �����Ŀ, d.�������, c.���, c.�걾����, c.ȡ��λ��, '��Ƭ' as �������, " & _
                " decode(b.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, b.��Ƭ�� as ����, " & _
                " decode(a.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') as ���״̬, decode(a.����״̬, 0, 'δ���', 1, '���ֽ��', '�ѽ��') as ����״̬,a.����״̬ as ����,d.����ʱ��, null as ִ�й��� " & _
                " from ����鵵��Ϣ a, ������Ƭ��Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ����ҽ����¼ e " & _
                " Where a.������Դ = 2 And a.��Ƭid = b.ID And b.�Ŀ�ID = c.�Ŀ�ID And c.����ҽ��id = d.����ҽ��id And d.ҽ��ID = e.ID and ����ID=[1] " & _
            " Union All " & _
                " select a.id as ��ԴID, a.����ҽ��ID, 3 as ������Դ, d.�����,e.����,e.�Ա�,e.����,e.ҽ������ as �����Ŀ, d.�������, c.���, c.�걾����, c.ȡ��λ��, " & _
                " decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
                " decode(b.�ؼ�ϸĿ,0,decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� || decode(b.��������,-1,'-��',0,'','-��' || b.��������) || ')' as ������ϸ, 1 as ����, " & _
                " decode(a.���״̬, 0, '�浵��', 1, '������ʧ', '����ʧ') as ���״̬, decode(a.����״̬, 0, 'δ���', 1, '���ֽ��', '�ѽ��') as ����״̬,a.����״̬ as ����,d.����ʱ��, null as ִ�й��� " & _
                " from ����鵵��Ϣ a, �����ؼ���Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ����ҽ����¼ e, ��������Ϣ f " & _
                " Where a.������Դ = 3 And a.�ؼ�id = b.ID And b.�Ŀ�ID = c.�Ŀ�ID And c.����ҽ��id = d.����ҽ��id And d.ҽ��ID = e.ID And b.����id = f.����id  and ����ID=[1]" & _
                " )order by ����"
    End If
    
'    If mblnMoved Then
'        strSql = strSql & " Union all " & GetMovedDataSql(strSql)
'    End If
    
    Set ufgArchivesDetail.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngArchivesId)
    Call ufgArchivesDetail.RefreshData
End Sub


Private Sub ReadArchivesInf(ByVal lngArchivesRowIndex As Long)
'��ȡ������Ϣ
    Dim strInf As String
    If lngArchivesRowIndex <= 0 Then Exit Sub
    
    strInf = "�������ƣ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "������ţ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_�������) & vbCrLf
    strInf = strInf & "�������ࣺ" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "��鷶Χ��" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��鷶Χ) & vbCrLf
    strInf = strInf & "���λ�ã�[����:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & "  ���:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_�������) & "  ����:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & "]" & vbCrLf
    strInf = strInf & "��ϸ��ַ��" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��ϸ��ַ) & vbCrLf
    strInf = strInf & "����˵����" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_����˵��) & vbCrLf
    strInf = strInf & "����״̬��" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_����״̬) & vbCrLf
    
    strInf = strInf & "��ʼ���ڣ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��ʼ����) & vbCrLf
    strInf = strInf & "�������ڣ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "�� �� �ˣ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_������) & vbCrLf
    strInf = strInf & "�������ڣ�" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_��������) & vbCrLf
    strInf = strInf & "�鵵ʱ�䣺" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_�鵵ʱ��)
    
    rtbDetail.Text = strInf
End Sub



Private Function Execute_ClearArchivesMaterial() As String
'��������������Ĳ�����Ϣ
    Dim i As Integer
    Dim strLog As String
    Dim blnAllowDel As Boolean
    
    strLog = ""
    For i = ufgArchivesDetail.GridRows - 1 To 1 Step -1
        If ufgArchivesDetail.GetRowCheck(i) Then
            blnAllowDel = True
            
            If ufgArchivesDetail.Text(i, gstrPatholCol_���״̬) <> "�浵��" Then
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "�����Ϊ [ " & ufgArchivesDetail.Text(i, gstrPatholCol_�����) & _
                                " ] �Ŀ��Ϊ [ " & ufgArchivesDetail.Text(i, gstrPatholCol_�Ŀ��) & "] ��" & _
                                ufgArchivesDetail.Text(i, gstrPatholCol_������ϸ) & ufgArchivesDetail.Text(i, gstrPatholCol_�������) & "�ѷ�����ʧ�����ܴӸõ������Ƴ���"
                                
                blnAllowDel = False
            End If
            
            If ufgArchivesDetail.Text(i, gstrPatholCol_����״̬) <> "δ���" And blnAllowDel Then
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "�����Ϊ [ " & ufgArchivesDetail.Text(i, gstrPatholCol_�����) & _
                                " ] �Ŀ��Ϊ [ " & ufgArchivesDetail.Text(i, gstrPatholCol_�Ŀ��) & "] ��" & _
                                ufgArchivesDetail.Text(i, gstrPatholCol_������ϸ) & ufgArchivesDetail.Text(i, gstrPatholCol_�������) & "�ѱ����ģ����ܴӸõ������Ƴ���"
                                
                blnAllowDel = False
            End If
        
            If blnAllowDel Then
                Call zlDatabase.ExecuteProcedure("ZL_������_�����뵵(" & ufgArchivesDetail.Text(i, gstrPatholCol_��ԴID) & ")", Me.Caption)
                
                '����ɾ���ɹ����Ƴ������е�����
                Call ufgArchivesDetail.RemoveRow(i)
            End If
        End If
    Next i
    
    Execute_ClearArchivesMaterial = strLog
End Function


Private Sub ConfigArchivesModifyState(ByVal blnIsEnterArchives As Boolean)
'���õ����޸�״̬
'blnIsEnterArchives���Ƿ�鵵(true���ѹ鵵, false��δ�鵵)
    Dim i As Long

    For i = 1 To tbrTools.Buttons.Count
        Select Case UCase(tbrTools.Buttons(i).Key)
            Case UCase("tbn_DelArchives"), UCase("tbn_UpdateArchives"), UCase("tbn_EnterArchives")
                tbrTools.Buttons(i).Enabled = Not blnIsEnterArchives
        End Select
    Next i
'
'    tabFilter.Item(0).Enabled = Not blnIsEnterArchives
    
'    If blnIsEnterArchives Then
'        tabFilter.Item(1).Selected = blnIsEnterArchives
'    Else
'        tabFilter.Item(0).Selected = Not blnIsEnterArchives
'    End If
    
    mnu_DelArchives.Enabled = Not blnIsEnterArchives
    mnu_UpdateArchives.Enabled = Not blnIsEnterArchives
    mnu_EnterArchives.Enabled = Not blnIsEnterArchives
    
    cmdDel.Enabled = Not blnIsEnterArchives
    cmdEnterArchives.Enabled = Not blnIsEnterArchives

End Sub

Private Sub ConfigArchivesPrintState(ByVal blnIsValidReport As Boolean)
'���ñ����ӡ��ť״̬
'blnIsValidReport:�Ƿ���Ч����0����Ч��1����Ч��
    cmdPreview.Enabled = blnIsValidReport
    cmdPrint.Enabled = blnIsValidReport
End Sub



Private Sub ufgArchives_OnColFormartChange()
On Error GoTo errHandle
    zlDatabase.SetPara "�����б�����", ufgArchives.GetColsString(ufgArchives), glngSys, G_LNG_PATHOLARCHIVES_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchives_OnColsNameReSet()
On Error GoTo errHandle
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Call QueryArchivesData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchives_OnSelChange()
On Error GoTo errHandle
    If ufgArchives.SelectionRow <= 0 Then
        mlngCurArchivesId = -1
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID) = "" Then Exit Sub
    
    '����ʱ���������ID��ͬ�������κδ���
    If mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID)) Then Exit Sub
    
    mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    Call SwitchArchivesFace(Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������)))
        
'    If tabFilter.Selected.Index = 1 Then Call LoadArchivesDetail(mlngCurArchivesId)

    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_����״̬) = "�ѹ鵵")
    
    Call ConfigArchivesPrintState(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������) <> "")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnColFormartChange()
'���������ϸ�б������
On Error GoTo errHandle
    If mblnIsFormLoaded Then
        zlDatabase.SetPara IIf(mcurMaterialType <> amtMaterial, "����ֽ����ϸ�б�����", "����������ϸ�б�����"), ufgArchivesDetail.GetColsString(ufgArchivesDetail), glngSys, G_LNG_PATHOLARCHIVES_NUM
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnColsNameReSet()
On Error GoTo errHandle

    If ufgArchivesDetail.DataGrid.Rows > 1 Then Call LoadArchivesDetail(mlngCurArchivesId)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnNewRow(ByVal Row As Long)
    '�жϲ��������Ƿ����ֲ��ϲŽ��� ����״̬���ж�
    If Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_��������)) = 1 Then
        If Nvl(ufgArchivesDetail.Text(Row, "����״̬")) <> "δ���" Then
            Call ufgArchivesDetail.DisableCheck(Row, ufgArchivesDetail.GetColIndexWithRowCheck)
        End If
    End If
End Sub

Private Sub ufgArchivesDetail_OnSelChange()
On Error GoTo errHandle
    If Not ufgArchivesDetail.IsSelectionRow Then Exit Sub
    
    Call LoadMaterialDetialNumber(ufgArchivesDetail.Text(ufgArchivesDetail.SelectionRow, gstrPatholCol_��ԴID))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadMaterialDetialNumber(ByVal lngMaterialArchivesId As Long)
'���������ϸ����
'lngMaterialArchivesId:���Ϲ鵵ID

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    'ֻ�е�������Ϊ��������ʱ���Ŷ�ȡ����������Ϣ
    If mcurMaterialType <> amtMaterial Then Exit Sub
    
    If Not txtNumberInf.Visible Then txtNumberInf.Visible = True
    
    strSql = "select a.ID, zl_�������_��ȡ����(a.ID) as �浵����,  nvl(b.��ʧ����, 0) as ��ʧ����, nvl(c.�ѽ�����, 0) as �ѽ�����  from ����鵵��Ϣ a, " & _
             " (select nvl(sum(��ʧ����),0) as ��ʧ����, �鵵ID from ������ʧ��Ϣ where �鵵ID=[1] group by �鵵ID) b, " & _
             " (select (nvl(sum(��������), 0) - nvl(sum(�黹����), 0)) as �ѽ�����, �鵵ID " & _
             " From ������Ĺ��� where  �黹״̬=0  and �鵵ID=[1] group by �鵵ID) c " & _
             " where a.id =b.�鵵ID(+) and a.id=c.�鵵ID(+) and a.id = [1]"
             
'    If mblnMoved Then
'        strSql = "select sum(�浵����) as �浵����, sum(��ʧ����) as ��ʧ����, sum(�ѽ�����) as �ѽ����� from (" & _
'                    strSql & " Union all" & GetMovedDataSql(strSql) & ") group by id"
'    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId)
    
    txtNumberInf.Text = "��ǰ����������0   �ڵ�������0   �ѽ�������0   ��ʧ������0"
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtNumberInf.Text = "��ǰ����������" & Nvl(rsData!�浵����) & _
                        "   �ڵ�������" & Val(Nvl(rsData!�浵����)) - Val(Nvl(rsData!��ʧ����)) - Val(Nvl(rsData!�ѽ�����)) & _
                        "   �ѽ�������" & Nvl(rsData!�ѽ�����) & _
                        "   ��ʧ������" & Nvl(rsData!��ʧ����)
End Sub




Private Sub ufgMaterialQuery_OnColFormartChange()
'������ϲ�ѯ�б������
On Error GoTo errHandle
    If mblnIsFormLoaded Then
        zlDatabase.SetPara IIf(mcurMaterialType <> amtMaterial, "����ֽ�ʲ�ѯ�б�����", "�������ϲ�ѯ�б�����"), ufgMaterialQuery.GetColsString(ufgMaterialQuery), glngSys, G_LNG_PATHOLARCHIVES_NUM
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

