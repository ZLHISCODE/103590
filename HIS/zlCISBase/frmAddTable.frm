VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.0#0"; "TTF16.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddTable 
   Caption         =   "���ӱ���"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmAddTable.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Tag             =   "�ɱ仯��"
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   3015
      Left            =   2760
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5318
      _0              =   $"frmAddTable.frx":1582
      _1              =   $"frmAddTable.frx":198B
      _2              =   $"frmAddTable.frx":1D94
      _3              =   $"frmAddTable.frx":219D
      _4              =   $"frmAddTable.frx":25A6
      _count          =   5
      _ver            =   2
   End
   Begin zl9CISBase.VisItem VisItem 
      Height          =   225
      Index           =   0
      Left            =   2160
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   397
      MousePointer    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowEdit       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7395
      TabIndex        =   3
      Top             =   720
      Width           =   7455
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   4620
         MaxLength       =   2
         TabIndex        =   11
         Top             =   20
         Width           =   450
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   6540
         MaxLength       =   2
         TabIndex        =   10
         Top             =   20
         Width           =   450
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   7
         Top             =   20
         Width           =   450
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   960
         MaxLength       =   2
         TabIndex        =   5
         Top             =   20
         Width           =   450
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   1
         Left            =   2970
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   20
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         OrigLeft        =   3705
         OrigTop         =   405
         OrigRight       =   3945
         OrigBottom      =   690
         Max             =   99
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   20
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         OrigLeft        =   3705
         OrigTop         =   405
         OrigRight       =   3945
         OrigBottom      =   690
         Max             =   99
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   3
         Left            =   6990
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         OrigLeft        =   3705
         OrigTop         =   405
         OrigRight       =   3945
         OrigBottom      =   690
         Max             =   99
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   2
         Left            =   5070
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         OrigLeft        =   3705
         OrigTop         =   405
         OrigRight       =   3945
         OrigBottom      =   690
         Max             =   99
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "�̶�����(&H)"
         Height          =   225
         Left            =   3600
         TabIndex        =   15
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "�̶�����(&L)"
         Height          =   210
         Left            =   5460
         TabIndex        =   14
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lblCol 
         Caption         =   "����(&C)"
         Height          =   210
         Left            =   1800
         TabIndex        =   6
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lblRow 
         Caption         =   "����(&R)"
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   75
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   915
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":2945
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":2B65
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":2D85
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":2FA5
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":31C5
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":33E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":35FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":3819
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":3A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":3C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":3E67
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":4081
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":477B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":4E75
            Key             =   "View"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":5091
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":52B1
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6465
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":54D1
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":56F1
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":5911
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":5B31
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":5D51
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":5F71
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":618B
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":63AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":65C5
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":67E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":69FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":6C19
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":7313
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":7A0D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":7C29
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTable.frx":7E49
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   11400
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   7200
      FixedBackground1=   0   'False
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ȡ��"
               Key             =   "ȡ��"
               Object.ToolTipText     =   "ȡ��"
               Object.Tag             =   "ȡ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ϲ�"
               Key             =   "�ϲ�"
               Object.ToolTipText     =   "�ϲ�"
               Object.Tag             =   "�ϲ�"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "��ɫ"
               Key             =   "��ɫ"
               Object.ToolTipText     =   "��ɫ"
               Object.Tag             =   "��ɫ"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˮƽ"
               Key             =   "ˮƽ"
               Object.ToolTipText     =   "ˮƽ����"
               Object.Tag             =   "ˮƽ"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ֱ"
               Key             =   "��ֱ"
               Object.ToolTipText     =   "��ֱ����"
               Object.Tag             =   "��ֱ"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "�鿴"
               Key             =   "�鿴"
               Object.ToolTipText     =   "���鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   14
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   16
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   195
      Left            =   1440
      TabIndex        =   17
      Top             =   7320
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7215
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      SimpleText      =   $"frmAddTable.frx":8069
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAddTable.frx":80B0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15055
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "ҳ������(&U)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "����(&N)"
         Enabled         =   0   'False
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "����(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPasteSpecial 
         Caption         =   "ѡ����ճ��(&S)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "ȡ��(&C)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignInsert 
         Caption         =   "����(&I)"
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "���Ԫ������(&I)"
            Index           =   0
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "���Ԫ������(&D)"
            Index           =   1
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "����(&R)"
            Index           =   2
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "����(&C)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Begin VB.Menu mnuDesignDeleteTable 
            Caption         =   "�Ҳ൥Ԫ������(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuDesignDeleteTable 
            Caption         =   "�·���Ԫ������(&U)"
            Index           =   1
         End
         Begin VB.Menu mnuDesignDeleteTable 
            Caption         =   "����(&R)"
            Index           =   2
         End
         Begin VB.Menu mnuDesignDeleteTable 
            Caption         =   "����(&C)"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuDesign 
      Caption         =   "��ʽ(&O)"
      Begin VB.Menu mnuFmtCell 
         Caption         =   "��Ԫ��(&E)"
      End
      Begin VB.Menu mnuFmtRow 
         Caption         =   "�и�(&R)"
      End
      Begin VB.Menu mnuFmtCol 
         Caption         =   "�п�(&C)"
      End
      Begin VB.Menu mnuDesign_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesign_Ass 
         Caption         =   "����������(&G)"
      End
      Begin VB.Menu mnuDesign_UnAss 
         Caption         =   "ȡ������(&U)"
      End
      Begin VB.Menu mnuDesign_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignMerge 
         Caption         =   "�ϲ���Ԫ(&M)"
      End
      Begin VB.Menu mnuDesignMergeCancel 
         Caption         =   "�����ϲ�(&Z)"
      End
      Begin VB.Menu mnuDesign_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignFont 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuDesignColor 
         Caption         =   "������ɫ(&C)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesignLineColor 
         Caption         =   "�����ɫ(&L)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesign_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignHsb 
         Caption         =   "ˮƽ����(&H)"
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "��߶���(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "���ж���(&C)"
            Index           =   1
         End
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "�ұ߶���(&R)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDesignVsb 
         Caption         =   "��ֱ����(&V)"
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "��������(&T)"
            Index           =   0
         End
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "���ж���(&C)"
            Index           =   1
         End
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "�ײ�����(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDesign_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignSize 
         Caption         =   "ͳһ�ߴ�(&S)"
         Begin VB.Menu mnuSize 
            Caption         =   "��ͬ�п�(&W)"
            Index           =   0
         End
         Begin VB.Menu mnuSize 
            Caption         =   "��ͬ�и�(&H)"
            Index           =   1
         End
         Begin VB.Menu mnuSize 
            Caption         =   "���߶���ͬ(&B)"
            Enabled         =   0   'False
            Index           =   2
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "����(&A)"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�2"
      Visible         =   0   'False
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "��߶���(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "���ж���(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "�ұ߶���(&R)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuShort3 
      Caption         =   "��ݲ˵�3"
      Visible         =   0   'False
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "��������(&T)"
         Index           =   0
      End
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "���ж���(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "�ײ�����(&B)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAddTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ģ�������õ��ľֲ�����˵��
Private mintColumn As Integer

Private mSelStartRow As Long
Private mSelEndRow As Long
Private mSelStartCol As Long
Private mSelEndCol As Long

Private mSvrMouseX As Long
Private mSvrMouseY As Long

Private mSvrRow As Long
Private mSvrCol As Long
Private OldTable As VBControlExtender
Attribute OldTable.VB_VarHelpID = -1

Private ShowTabs As Integer, TableEnabled As Boolean, ShowRowHeading As Boolean, ShowColHeading As Boolean
Private hOldWnd As Long
Private bNotRunSelChange As Boolean

Public WithEvents theTable As VBControlExtender
Attribute theTable.VB_VarHelpID = -1
Public theTableID As String, TableTitle As String

Private Sub Form_Activate()
    If Me.Tag = "Loading" And Len(theTableID) > 0 Then
        Me.Tag = ""
        
        Me.MousePointer = vbHourglass
        BeginShowProgress
        theTable.Tag = theTableID
        ReadTable theTable, theTableID, , Me.prbRefresh
        Me.prbRefresh.Visible = False
        Me.MousePointer = vbDefault
    
    
        theTable.SheetName(1) = TableTitle
        theTable.Modified = False
        '������ʾǰ�����ݳ�ʼ������
        txt(1).Text = theTable.MaxRow
        txt(2).Text = theTable.MaxCol
        txt(3).Text = theTable.FixedRows
        txt(4).Text = theTable.FixedCols
    
        udn(0).Value = txt(1).Text
        udn(1).Value = txt(2).Text
        udn(2).Value = txt(3).Text
        udn(3).Value = txt(4).Text
    End If
    If hOldWnd > 0 Then
        Me.MousePointer = vbHourglass
        BeginShowProgress
        RefreshObject False, Me.prbRefresh
        Me.prbRefresh.Visible = False
        Me.MousePointer = vbDefault
    End If
    theTable.SetFocus
End Sub

Private Sub Form_Load()
    Dim cellFormat As TTF160Ctl.F1CellFormat
    
    Call RestoreWinState(Me, App.ProductName)
    
    If theTable Is Nothing Then
        hOldWnd = 0
'        Set theTable = Me.Controls.Add("ttf16.ttf1.6", "theTable", Me)
        Set theTable = F1Book1
        InitTable theTable
        
        Me.Tag = "Loading" 'Ҫ��ȡ���
    Else
        hOldWnd = GetParent(theTable.hwnd)
        
        SetParent theTable.hwnd, Me.hwnd
        With theTable
            ShowTabs = .ShowTabs
            TableEnabled = .Enabled
            ShowColHeading = .ShowColHeading: ShowRowHeading = .ShowRowHeading
            .ShowTabs = F1TabsBottom
            .Enabled = True
            .ShowColHeading = True: .ShowRowHeading = True
        End With
    End If
    
    theTable.SheetName(1) = TableTitle
    theTable.Modified = False
    '������ʾǰ�����ݳ�ʼ������
    txt(1).Text = theTable.MaxRow
    txt(2).Text = theTable.MaxCol
    txt(3).Text = theTable.FixedRows
    txt(4).Text = theTable.FixedCols

    udn(0).Value = txt(1).Text
    udn(1).Value = txt(2).Text
    udn(2).Value = txt(3).Text
    udn(3).Value = txt(4).Text
End Sub

Private Sub Form_Resize()
    '���ݴ���״̬,���������и��ؼ�����ʾλ��
    On Error Resume Next
    With stbThis
        .Align = vbAlignNone
        .Top = Me.ScaleHeight - .Height: .Width = Me.ScaleWidth
        .Align = vbAlignBottom
    End With
    With cbrThis
        .Align = vbAlignNone
        .Width = Me.ScaleWidth
        .Align = vbAlignTop
    End With
    
    With Picture1
        .Top = cbrThis.Top + IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth
    End With
    With theTable
        .Left = 0
        .Top = Picture1.Top + Picture1.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
        .Visible = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)

    If theTable.Modified And mnuEditSave.Visible Then
        If MsgBox("������޸ģ��Ƿ񱣴棿", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then mnuEditSave_Click
    End If
    
    If hOldWnd > 0 Then
        WriteToTable
        SetParent theTable.hwnd, hOldWnd
        theTable.ShowTabs = ShowTabs
        theTable.Enabled = TableEnabled
        theTable.ShowRowHeading = ShowRowHeading
        theTable.ShowColHeading = ShowColHeading
    Else
        Set theTable = Nothing
    End If
End Sub

Private Sub mnuDesign_Ass_Click()
    Dim sItemID As String
    Dim i As Long, j As Long
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    
    With theTable
        iStartRow = .SelStartRow: iEndRow = IIf(.SelEndRow = 65536, .MaxRow, .SelEndRow)
        iStartCol = .SelStartCol: iEndCol = IIf(.SelEndCol = 256, .MaxCol, .SelEndCol)
    End With
    
    frmSelVis.ItemID = ""
    frmSelVis.Show vbModal, Me: DoEvents
    sItemID = frmSelVis.ItemID
    If Len(sItemID) = 0 Then Exit Sub
    
    With theTable
        For i = iStartRow To iEndRow
            For j = iStartCol To iEndCol
                AddObject theTable, i, j, sItemID, , , Me
            Next j
        Next i
    End With
End Sub

Private Sub mnuDesign_UnAss_Click()
    Dim i As Long, j As Long
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    With theTable
        iStartRow = .SelStartRow: iEndRow = IIf(.SelEndRow = 65536, .MaxRow, .SelEndRow)
        iStartCol = .SelStartCol: iEndCol = IIf(.SelEndCol = 256, .MaxCol, .SelEndCol)
        
        For i = iStartRow To iEndRow
            For j = iStartCol To iEndCol
                RemoveObject theTable, i, j, Me
            Next j
        Next i
    End With
End Sub

Private Sub mnuDesignColor_Click()
    '����ָ����Ԫ���������ɫ,����һ��ָ�������Ԫ��
End Sub

Private Sub mnuDesignDeleteTable_Click(Index As Integer)
    On Error Resume Next
    With theTable
        Select Case Index
            Case 0
                .DeleteRange .SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol, F1ShiftHorizontal
            Case 1
                .DeleteRange .SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol, F1ShiftVertical
            Case 2
                .DeleteRange .SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol, F1ShiftRows
            Case 3
                .DeleteRange .SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol, F1ShiftCols
        End Select
    End With
    
    Me.MousePointer = vbHourglass
    BeginShowProgress
    RefreshObject , Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuDesignFont_Click()
    theTable.FormatCellsDlg F1FontPage
End Sub

Private Sub mnuDesignInsertTable_Click(Index As Integer)
    On Error Resume Next
    With theTable
        Select Case Index
            Case 0
                .InsertRange .Row, .Col, .Row, .Col, F1ShiftHorizontal
            Case 1
                .InsertRange .Row, .Col, .Row, .Col, F1ShiftVertical
            Case 2
                .InsertRange .Row, .Col, .Row, .Col, F1ShiftRows
            Case 3
                .InsertRange .Row, .Col, .Row, .Col, F1ShiftCols
        End Select
    End With
    
    Me.MousePointer = vbHourglass
    BeginShowProgress
    RefreshObject , Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuDesignLineColor_Click()
    '���ñ�����������ɫ
End Sub

Private Sub mnuDesignMerge_Click()
    Dim cellFormat As F1CellFormat
    
    On Error Resume Next
    Set cellFormat = theTable.GetCellFormat
    cellFormat.MergeCells = True
    theTable.SetCellFormat cellFormat
    
    Me.MousePointer = vbHourglass
    BeginShowProgress
    RefreshObject , Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuDesignMergeCancel_Click()
    '�����ϲ���Ԫ��
    Dim cellFormat As F1CellFormat
    Dim iRow As Long, iCol As Long
    Dim strItemInfo As String
    
    On Error Resume Next
    iRow = theTable.SelStartRow: iCol = theTable.SelStartCol
    Set cellFormat = theTable.GetCellFormat
    strItemInfo = cellFormat.ValidationText
    cellFormat.ValidationText = ""
    cellFormat.MergeCells = False
    theTable.SetCellFormat cellFormat
    
    If Len(strItemInfo) > 0 Then
        theTable.SetSelection theTable.SelEndRow, theTable.SelEndCol, theTable.SelEndRow, theTable.SelEndCol
        theTable.SetSelection iRow, iCol, iRow, iCol
        
        Set cellFormat = theTable.GetCellFormat
        cellFormat.ValidationText = strItemInfo
        theTable.SetCellFormat cellFormat
    End If

    Me.MousePointer = vbHourglass
    BeginShowProgress
    RefreshObject , Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditCancel_Click()
    'ȡ���Ա����޸Ļ�����
    'picLvwBack.Tag=1��ʾ�������;picLvwBack.Tag=2��ʾ�޸ı��
    
'    If bEdit = True Then
'        If MsgBox("�޸ĺ�ı��Ҫ�������Ч��ȷ�ϲ�������˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    End If
'
'    bEdit = False
'    Call Reset
'    If picLvwBack.Tag <> "" And Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
'
'    picLvwBack.Tag = ""
'    picLvwBack.Enabled = True
'    picEdit.Enabled = False
'
'    Call AdjustEnabled
    
End Sub

Private Sub mnuEditCopy_Click()
    theTable.EditCopy
End Sub

Private Sub mnuEditCut_Click()
    theTable.EditCut
End Sub

Private Sub mnuEditSave_Click()
    '������Ԫ�ؼ���������
    Me.MousePointer = vbHourglass
    BeginShowProgress
    gcnOracle.Execute "Delete From ���������� Where Ԫ��ID=" & theTable.Tag
    SaveTable theTable, , Me, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    theTable.Modified = False
End Sub

Private Sub mnuEditSelectAll_Click()
    theTable.SetSelection 1, 1, theTable.MaxRow, theTable.MaxCol
End Sub

Private Sub mnuFileExcel_Click()
'    Call PrintObject(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
'    Call PrintObject(2)
End Sub

Private Sub mnuFilePrint_Click()
'    Call PrintObject(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFmtCell_Click()
    theTable.FormatCellsDlg F1AllPages 'F1AlignmentPage + F1FontPage '+ F1GeneralPage + F1OptionsPage  '+ F1EditPage
End Sub

Private Sub mnuFmtCol_Click()
    theTable.ColWidthDlg
End Sub

Private Sub mnuFmtRow_Click()
    theTable.RowHeightDlg
End Sub

Private Sub mnuhelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuHsbAlign_Click(Index As Integer)
    Dim cellFormat As F1CellFormat
    
    With theTable
        Set cellFormat = .GetCellFormat
        cellFormat.AlignHorizontal = IIf(Index = 0, F1HAlignLeft, IIf(Index = 1, F1HAlignCenter, F1HAlignRight))
        .SetCellFormat cellFormat
    End With
End Sub

Private Sub mnuPaste_Click()
    On Error GoTo PasteErr
    theTable.EditPaste
    Exit Sub
PasteErr:
    MsgBox theTable.ErrorNumberToText(Err), vbExclamation, gstrSysName
End Sub

Private Sub mnuPasteSpecial_Click()
    theTable.PasteSpecialDlg
End Sub

Private Sub mnuShort2Hsb_Click(Index As Integer)
    Call mnuHsbAlign_Click(Index)
End Sub

Private Sub mnuShort3Vsb_Click(Index As Integer)
    Call mnuVsbAlign_Click(Index)
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim i As Long
    
    With theTable
        Select Case Index
            Case 0          '��ͬ�п�
                .ColWidthDlg
                For i = .SelStartCol To IIf(.SelEndCol = 256, .MaxCol, .SelEndCol)
                    .ColWidth(i) = .ColWidth(.Col)
                Next
            Case 1          '��ͬ�и�
                .RowHeightDlg
                For i = .SelStartRow To IIf(.SelEndRow = 65536, .MaxRow, .SelEndRow)
                    .RowHeight(i) = .RowHeight(.Row)
                Next
            End Select
    End With
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub


Private Sub mnuViewToolText_Click()
    Dim i As Long

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MINHEIGHT = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mnuVsbAlign_Click(Index As Integer)
    Dim cellFormat As F1CellFormat
    
    With theTable
        Set cellFormat = .GetCellFormat
        cellFormat.AlignVertical = IIf(Index = 0, F1VAlignTop, IIf(Index = 1, F1VAlignCenter, F1VAlignBottom))
        .SetCellFormat cellFormat
    End With
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePreview_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "����"
    Case "�޸�"
    Case "ɾ��"
    Case "�鿴"
'        If lvw.View < 3 Then
'            Call mnuViewIcon_Click(lvw.View + 1)
'        Else
'            Call mnuViewIcon_Click(0)
'        End If
    Case "����"
        Call mnuHelpTopic_Click
    Case "����"
        Call mnuEditSave_Click
    Case "ȡ��"
        Call mnuEditCancel_Click
    Case "�ϲ�"
        Call mnuDesignMerge_Click
    Case "����"
        Call mnuDesignMergeCancel_Click
    Case "����"
        Call mnuDesignFont_Click
    Case "��ɫ"
        Call mnuDesignColor_Click
    Case "ˮƽ"
        Me.PopupMenu mnuShort2
    Case "��ֱ"
        Me.PopupMenu mnuShort3
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub theTable_GotFocus()
    bNotRunSelChange = False
End Sub

Private Sub theTable_LostFocus()
    bNotRunSelChange = True
End Sub

Private Sub theTable_ObjectEvent(Info As EventInfo)
    Dim iDecPos As Integer
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
    
    Select Case LCase(Info.Name)
        Case "dblclick"
            theTable.StartEdit False, True, False
        Case "objgotfocus"
'            With theTable
'                .Row = .ObjCellRow(Info.EventParameters(1)): .Col = .ObjCellCol(Info.EventParameters(1))
'            End With
        Case "endedit"
            Dim EditString As String
            
            EditString = Info.EventParameters("EditString").Value
            With theTable
                If IsNumeric(EditString) Then
                    iDecPos = InStr(EditString, ".")
                    If iDecPos > 0 And iDecPos < Len(EditString) Then
                        .NumberFormat = "#." + String(Len(EditString) - iDecPos, "0")
                    Else
                        .NumberFormat = "General"
                    End If
                Else
                    .NumberFormat = "General"
                End If
                .TextRC(.Row, .Col) = EditString
                .SetRowHeightAuto .Row, 1, .Row, .MaxCol, True
            End With
            bNotRunSelChange = False
        Case "canceledit"
            bNotRunSelChange = False
        Case "topleftchanged"
            '���û����������ģ�������
            If bNotRunSelChange Then Exit Sub
            
            bNotRunSelChange = True
            Proc_Table_TopLeftChanged theTable, Me
            bNotRunSelChange = False
        Case "selchange"
            On Error Resume Next
            '���û����������ģ�������
            If bNotRunSelChange Then Exit Sub
            If Not Me.Visible Or Me.ActiveControl.Name <> "theTable" Then Exit Sub
            With theTable
                Set objCellFormat = .GetCellFormat
                If Len(objCellFormat.ValidationText) > 0 Then
                    aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                    Me.VisItem(aVisItemInfo(1)).SetFocus
                End If
            End With
        Case "keypress"
            On Error Resume Next
            With theTable
                Set objCellFormat = .GetCellFormat
                If Len(objCellFormat.ValidationText) > 0 Then
                    Info.EventParameters("KeyAscii").Value = 0
                End If
            End With
        Case "mouseup"
            If Info.EventParameters(0).Value = 2 Then Call PopupMenu(Me.mnuDesign, 2)
        Case "startedit"
            On Error Resume Next
            bNotRunSelChange = True
            With theTable
                Set objCellFormat = .GetCellFormat
                If Len(objCellFormat.ValidationText) > 0 Then
                    Info.EventParameters(1).Value = True
                End If
            End With
    End Select
End Sub
'������Ĺ������¼�
Private Sub Proc_Table_TopLeftChanged(theTable As TTF160Ctl.F1Book, Optional objParent As Object)
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim tmpCtrl As Control, aCellRC() As String
    Dim bValidCtrl As Boolean
    Dim frmParent As Object
        
    On Error Resume Next
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '������
    Else
        Set frmParent = objParent
    End If
    With theTable
        iCurrRow = .Row: iCurrCol = .Col
        iStartRow = .SelStartRow: iEndRow = .SelEndRow
        iStartCol = .SelStartCol: iEndCol = .SelEndCol

        .SetSelection iStartRow, iStartCol, iStartRow, iStartCol
        For Each tmpCtrl In frmParent.Controls
            bValidCtrl = True
            If Not (tmpCtrl.Name = "VisItem" And Len(tmpCtrl.Tag) > 0) Then bValidCtrl = False
            
            If bValidCtrl Then
                aCellRC = Split(tmpCtrl.Tag, ",")
                .SetActiveCell aCellRC(0), aCellRC(1)
    
                tmpCtrl.Visible = False
                '��Ԫ�ɼ�
                If .RangeShown(.SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol) = 1 Then
                    Set objRect = .RangeToTwipsEx(.SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol)
        
                    tmpCtrl.Left = objRect.Left + .Left + 30
                    tmpCtrl.Top = objRect.Top + .Top + 30
                    tmpCtrl.Width = objRect.Width - 30
                    tmpCtrl.Height = objRect.Height - 30
                    If objRect.Width - 30 < tmpCtrl.Width Then
                        .ColWidthTwips(.SelStartCol) = _
                            .ColWidthTwips(.SelStartCol) + tmpCtrl.Width - (objRect.Width - 30)
                    End If
                    If objRect.Height - 30 < tmpCtrl.Height Then
                        .RowHeight(.SelStartRow) = _
                            .RowHeight(.SelStartRow) + tmpCtrl.Height - (objRect.Height - 30)
                    End If
                    tmpCtrl.Visible = True
'                    If tmpCtrl.Left < .Left Or tmpCtrl.Left + tmpCtrl.Width > .Left + .Width Or _
'                        tmpCtrl.Top < .Top Or tmpCtrl.Top + tmpCtrl.Height > .Top + .Height Then
'                        tmpCtrl.Visible = False
'                    Else
'                        tmpCtrl.Visible = True
'                    End If
                End If
            End If
        Next
        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .SetActiveCell iCurrRow, iCurrCol
    End With
End Sub
'����ˢ��������
Private Sub RefreshObject(Optional ByVal HasVisItem As Boolean = True, Optional objProgBar As ProgressBar)
    Dim iDecPos As Integer
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
    
    On Error Resume Next
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    For Each tmpCtrl In Me.Controls
        If tmpCtrl.Name = "VisItem" Then tmpCtrl.Visible = False
    Next
        
    objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = theTable.MaxRow * theTable.MaxCol
    For iRow = 1 To theTable.MaxRow
        For iCol = 1 To theTable.MaxCol
            theTable.SetActiveCell iRow, iCol

            Set objCellFormat = theTable.GetCellFormat
            If Len(objCellFormat.ValidationText) > 0 And iRow = theTable.SelStartRow And iCol = theTable.SelStartCol Then
                aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                
                objCellFormat.ValidationText = ""
                theTable.SetCellFormat objCellFormat
                
                If Not HasVisItem Then
                    AddObject theTable, iRow, iCol, CLng(aVisItemInfo(0)), True, theTable.TextRC(iRow, iCol), Me
                Else
                    AddObject theTable, iRow, iCol, CLng(aVisItemInfo(0)), True, Me.VisItem(aVisItemInfo(1)).Value, Me
                End If
            End If
                
            objProgBar.Value = (iRow - 1) * theTable.MaxCol + iCol
        Next iCol
    Next iRow
    For Each tmpCtrl In Me.Controls
        If tmpCtrl.Name = "VisItem" And Not tmpCtrl.Visible Then Unload tmpCtrl
    Next
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub
'���������ֵд�뵥Ԫ����
Private Sub WriteToTable()
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
    
    On Error Resume Next
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    For iRow = 1 To theTable.MaxRow
        For iCol = 1 To theTable.MaxCol
            theTable.SetActiveCell iRow, iCol

            Set objCellFormat = theTable.GetCellFormat
            If Len(objCellFormat.ValidationText) > 0 And iRow = theTable.SelStartRow And iCol = theTable.SelStartCol Then
                aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                theTable.TextRC(iRow, iCol) = Me.VisItem(aVisItemInfo(1)).Value
            End If
        Next iCol
    Next iRow
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub
Private Sub txt_GotFocus(Index As Integer)
    With txt(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Select Case Index
            Case 1
                txt(2).SetFocus
            Case 2
                txt_LostFocus (2)
                theTable.SetFocus
            Case 3
                txt(4).SetFocus
            Case 4
                txt_LostFocus (4)
                theTable.SetFocus
        End Select
        Exit Sub
    End If
    
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or ifEditKey(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Not IsNumeric(txt(Index).Text) Then txt(Index).Text = 0
    If Index < 3 And Val(txt(Index).Text) <= 0 Then
        MsgBox "����������Ϊ0��������Ҫһ�к�һ�У�", vbExclamation + vbOKOnly, gstrSysName
        Select Case Index
            Case 1
                txt(1).Text = theTable.MaxRow
            Case 2
                txt(2).Text = theTable.MaxCol
        End Select
        txt(Index).SetFocus
        Exit Sub
    End If
    If Index > 2 And (Val(txt(Index).Text) <= -1 Or Val(txt(3).Text) > theTable.MaxRow - 1 Or Val(txt(4).Text) > theTable.MaxCol - 1) Then
        If Val(txt(Index).Text) <= -1 Then MsgBox "�̶�����������Ϊ������", vbExclamation + vbOKOnly, gstrSysName
        If Val(txt(3).Text) > theTable.MaxRow Then MsgBox "�̶��������ܳ�����������", vbExclamation + vbOKOnly, gstrSysName
        If Val(txt(4).Text) > theTable.MinCol Then MsgBox "�̶��������ܳ�����������", vbExclamation + vbOKOnly, gstrSysName
        Select Case Index
            Case 3
                txt(3).Text = theTable.FixedRows
            Case 4
                txt(4).Text = theTable.FixedCols
        End Select
        txt(Index).SetFocus
        Exit Sub
    End If

    Select Case Index
        Case 1
            If Val(txt(1).Text) <> theTable.MaxRow Then
                theTable.MaxRow = Val(txt(1).Text)
                If Me.Visible Then
                    Me.MousePointer = vbHourglass
                    BeginShowProgress
                    RefreshObject , Me.prbRefresh
                    Me.prbRefresh.Visible = False
                    Me.MousePointer = vbDefault
                End If
            End If
            udn(0).Value = txt(1).Text
            If theTable.MaxRow <= theTable.FixedRows Then
                theTable.FixedRows = theTable.MaxRow - 1
                txt(3).Text = theTable.FixedRows
                udn(2).Value = txt(3).Text
            End If
        Case 2
            If Val(txt(2).Text) <> theTable.MaxCol Then
                theTable.MaxCol = Val(txt(2).Text)
                If Me.Visible Then
                    Me.MousePointer = vbHourglass
                    BeginShowProgress
                    RefreshObject , Me.prbRefresh
                    Me.prbRefresh.Visible = False
                    Me.MousePointer = vbDefault
                End If
            End If
            udn(1).Value = txt(2).Text
            If theTable.MaxCol <= theTable.FixedCols Then
                theTable.FixedCols = theTable.MaxCol - 1
                txt(4).Text = theTable.FixedCols
                udn(3).Value = txt(4).Text
            End If
        Case 3
            If Val(txt(3).Text) <> theTable.FixedRows Then theTable.FixedRows = Val(txt(3).Text)
            udn(2).Value = txt(3).Text
        Case 4
            If Val(txt(4).Text) <> theTable.FixedCols Then theTable.FixedCols = Val(txt(4).Text)
            udn(3).Value = txt(4).Text
    End Select
    
    theTable.SetActiveCell 1, 1
End Sub

Private Sub udn_Change(Index As Integer)
    Dim lngOldRows As Long, lngOldCols As Long
    Dim lngCurrRow As Long, lngCurrCol As Long
    Dim cellFormat  As TTF160Ctl.F1CellFormat
    
    On Error Resume Next
    Select Case Index
        Case 0
            If theTable.MaxRow = udn(Index).Value Then Exit Sub
            lngOldRows = theTable.MaxRow
            theTable.MaxRow = udn(Index).Value
            
            If lngOldRows < theTable.MaxRow Then
                '���������Ԫ������
                With theTable
                    .SetSelection lngOldRows + 1, 1, .MaxRow, .MaxCol
                    Set cellFormat = .GetCellFormat
                    cellFormat.ProtectionLocked = False
                    cellFormat.MergeCells = False
                    .SetCellFormat cellFormat
                    .SetSelection 1, 1, 1, 1
                End With
            End If
            
            If Me.Visible Then
                Me.MousePointer = vbHourglass
                BeginShowProgress
                RefreshObject , Me.prbRefresh
                Me.prbRefresh.Visible = False
                Me.MousePointer = vbDefault
            End If
            txt(1).Text = udn(Index).Value
            
            If Not IsNumeric(txt(3).Text) Then txt(3).Text = "0"
            If theTable.MaxRow <= Val(txt(3).Text) Then
                theTable.FixedRows = theTable.MaxRow - 1
                txt(3).Text = theTable.FixedRows
                udn(2).Value = txt(3).Text
            End If
            theTable.SetActiveCell 1, 1
        Case 1
            If theTable.MaxCol = udn(Index).Value Then Exit Sub
            lngOldCols = theTable.MaxCol
            theTable.MaxCol = udn(Index).Value
            
            If lngOldCols < theTable.MaxCol Then
                '���������Ԫ������
                With theTable
                    .SetSelection 1, lngOldCols + 1, .MaxRow, .MaxCol
                    Set cellFormat = .GetCellFormat
                    cellFormat.ProtectionLocked = False
                    cellFormat.MergeCells = False
                    .SetCellFormat cellFormat
                    .SetSelection 1, 1, 1, 1
                End With
            End If
            
            If Me.Visible Then
                Me.MousePointer = vbHourglass
                BeginShowProgress
                RefreshObject , Me.prbRefresh
                Me.prbRefresh.Visible = False
                Me.MousePointer = vbDefault
            End If
            txt(2).Text = udn(Index).Value
            
            If Not IsNumeric(txt(4).Text) Then txt(4).Text = "0"
            If theTable.MaxCol <= Val(txt(4).Text) Then
                theTable.FixedCols = theTable.MaxCol - 1
                txt(4).Text = theTable.FixedCols
                udn(3).Value = txt(4).Text
            End If
            theTable.SetActiveCell 1, 1
        Case 2
            If udn(Index).Value >= theTable.MaxRow Then
                udn(Index).Value = theTable.MaxRow - 1
            End If
            theTable.FixedRows = udn(Index).Value
            txt(3).Text = udn(Index).Value
            theTable.SetActiveCell theTable.FixedRows + 1, theTable.FixedCols + 1
        Case 3
            If udn(Index).Value >= theTable.MaxCol Then
                udn(Index).Value = theTable.MaxCol - 1
            End If
            theTable.FixedCols = udn(Index).Value
            txt(4).Text = udn(Index).Value
            theTable.SetActiveCell theTable.FixedRows + 1, theTable.FixedCols + 1
    End Select
    theTable.ShowActiveCell
End Sub

'-----------------------------------------------------------------------------------------------------------------
'
'�������Զ��庯������̲���,������ģ����ʹ��
'
'-----------------------------------------------------------------------------------------------------------------
Private Sub ModulePrivs()
    '����ģ��Ȩ��,������������ػ���ʾ
    'Ȩ����:��ɾ��
    
    mnuEdit.Visible = True
    mnuDesign.Visible = True
        
    If InStr(gstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuDesign.Visible = False
        
        tbrThis.Buttons("����").Visible = False
        tbrThis.Buttons("�޸�").Visible = False
        tbrThis.Buttons("ɾ��").Visible = False
        tbrThis.Buttons("Split_2").Visible = False
        tbrThis.Buttons("����").Visible = False
        tbrThis.Buttons("ȡ��").Visible = False
        tbrThis.Buttons("Split_3").Visible = False
        
        tbrThis.Buttons("�ϲ�").Visible = False
        tbrThis.Buttons("����").Visible = False
        tbrThis.Buttons("����").Visible = False
        tbrThis.Buttons("��ɫ").Visible = False
        tbrThis.Buttons("ˮƽ").Visible = False
        tbrThis.Buttons("��ֱ").Visible = False
        tbrThis.Buttons("Split_4").Visible = False
    End If
End Sub

Private Sub ExChange(x As Long, y As Long)
    '����X��Y��ֵ
    Dim Tmp As Long
    
    Tmp = x
    x = y
    y = Tmp
End Sub

Private Sub VisItem_GotFocus(Index As Integer)
    Dim aCellInfo() As String

    On Error Resume Next
    aCellInfo = Split(VisItem(Index).Tag, ",")
    
    theTable.SetActiveCell aCellInfo(0), aCellInfo(1)
End Sub

Private Sub VisItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim aCellInfo() As String
    
    On Error Resume Next
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        theTable.SetFocus
        zlcommfun.PressKey CByte(KeyCode)
    End If
End Sub

Private Sub BeginShowProgress()
    With prbRefresh
        .Left = stbThis.Panels(2).Left + 50
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width - 50
        .Visible = stbThis.Visible
    End With
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

