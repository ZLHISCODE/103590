VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMdi 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9990
   Icon            =   "frmMdi.frx":0000
   KeyPreview      =   -1  'True
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6570
   ScaleWidth      =   9990
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateConnect 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImgUsualBlack 
      Left            =   30
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgUsualColor 
      Left            =   600
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicBackBitmap 
      AutoRedraw      =   -1  'True
      Height          =   585
      Left            =   60
      Picture         =   "frmMdi.frx":1CFA
      ScaleHeight     =   525
      ScaleWidth      =   1605
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   30
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":DEB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E0CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E2E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E600
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E91A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":EC34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F32E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F548
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F762
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   600
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":107F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":10A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":10C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":114CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1191E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":12018
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":12232
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1244C
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9990
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   5295
      MinHeight1      =   720
      Width1          =   1425
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "TbrUsual"
      MinWidth2       =   1200
      MinHeight2      =   330
      Width2          =   675
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Begin MSComctlLib.Toolbar TbrUsual 
         Height          =   330
         Left            =   8700
         TabIndex        =   7
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgUsualBlack"
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   1270
         ButtonWidth     =   1455
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBlack"
         HotImageList    =   "ImgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "printbar"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ֵ����"
               Key             =   "Dictionary"
               Object.ToolTipText     =   "�ֵ����"
               Object.Tag             =   "�ֵ����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ϣ�շ�"
               Key             =   "Message"
               Object.ToolTipText     =   "��Ϣ�շ�"
               Object.Tag             =   "��Ϣ�շ�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ϵͳѡ��"
               Key             =   "Choose"
               Object.ToolTipText     =   "����ѡ��"
               Object.Tag             =   "����ѡ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bar"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ⲿ��"
               Key             =   "Check"
               Object.ToolTipText     =   "��ⲿ��"
               Object.Tag             =   "��ⲿ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӹ���"
               Key             =   "����"
               Object.ToolTipText     =   "��ӹ�������"
               Object.Tag             =   "����"
               ImageKey        =   "Tool"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4770
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   5475
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9657
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8421504
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TvwMenu 
      Height          =   2745
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4842
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   1050
      Left            =   285
      TabIndex        =   0
      Top             =   4845
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1852
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   6216
      Width           =   9984
      _ExtentX        =   17621
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   1764
            Picture         =   "frmMdi.frx":134DE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11774
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
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
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Image ImgTry 
      Height          =   675
      Left            =   4020
      Top             =   3060
      Width           =   1125
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReLogin 
         Caption         =   "ע��(&R)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuOper 
      Caption         =   "����(&O)"
      Begin VB.Menu mnuOperDefault 
         Caption         =   "Default"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuOrderMenu 
         Caption         =   "�������й��ܲ˵�(&L)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolTester 
         Caption         =   "ʹ��SQL�ٶȲ��Թ���(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolIndividuation 
         Caption         =   "ʹ�ø��Ի�����(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuToolNotify 
         Caption         =   "��Ϣ֪ͨ(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolShowDisReport 
         Caption         =   "��ʾͣ�ñ���(&P)"
      End
      Begin VB.Menu mnuToolSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolDictonary 
         Caption         =   "�ֵ������(&D)"
      End
      Begin VB.Menu mnuToolMessage 
         Caption         =   "��Ϣ�շ�����(&M)"
      End
      Begin VB.Menu mnuToolNotice 
         Caption         =   "������Ϣ����(&R)"
      End
      Begin VB.Menu mnuTooleSelect 
         Caption         =   "ϵͳѡ��(&S)"
      End
      Begin VB.Menu mnuToolExcel 
         Caption         =   "����&EXCEL����"
      End
      Begin VB.Menu mnuToolSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolHistory 
         Caption         =   "�����ʷ��¼(&H)"
      End
      Begin VB.Menu mnuToolOutTool 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOutToolSet 
         Caption         =   "��ӹ�������(&O)"
      End
      Begin VB.Menu mnuToolOutToolExecute 
         Caption         =   "����(&1)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRepair 
      Caption         =   "�޸�(&R)"
      Begin VB.Menu mnuRepairIndividuationClear 
         Caption         =   "������������쳣(&C)"
      End
      Begin VB.Menu mnuRepairComponent 
         Caption         =   "��ⰲװ����(&T)"
      End
      Begin VB.Menu mnuRepairClientUpdate 
         Caption         =   "�ͻ����޸�(&U)"
      End
   End
   Begin VB.Menu History 
      Caption         =   "��ʷ(&O)"
      Begin VB.Menu HistoryItem 
         Caption         =   "��"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      Begin VB.Menu mnuWindowList 
         Caption         =   "���Ŵ���(&L)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
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
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu MnuRightMenu 
      Caption         =   "�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu MnuRightAbout 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu MnuRightBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightTester 
         Caption         =   "ʹ��SQL�ٶȲ��Թ���(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRightIndividuation 
         Caption         =   "ʹ�ø��Ի�����(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRightNotify 
         Caption         =   "��Ϣ֪ͨ(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRightShowDisReport 
         Caption         =   "��ʾͣ�ñ���(&P)"
      End
      Begin VB.Menu MnuRightBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightDictonary 
         Caption         =   "�ֵ������(&D)"
      End
      Begin VB.Menu mnuRightMessage 
         Caption         =   "��Ϣ�շ�����(&M)"
      End
      Begin VB.Menu mnuRightNotice 
         Caption         =   "������Ϣ����(&T)"
      End
      Begin VB.Menu MnuRightStyle 
         Caption         =   "ϵͳѡ��(&S)"
      End
      Begin VB.Menu MnuRightExcel 
         Caption         =   "����&EXCEL����"
      End
      Begin VB.Menu MnuRightBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightIndividuationClear 
         Caption         =   "������������쳣(&C)"
      End
      Begin VB.Menu MnuRightComponent 
         Caption         =   "��ⰲװ����(&C)"
      End
      Begin VB.Menu mnuRightClientUpdate 
         Caption         =   "�ͻ����޸�(&U)"
      End
      Begin VB.Menu MnuRightHistory 
         Caption         =   "�����ʷ��¼(&H)"
      End
      Begin VB.Menu MnuRightBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRightSetColor 
         Caption         =   "����������ɫ(&O)"
      End
      Begin VB.Menu MnuRightBackBmp 
         Caption         =   "ѡ�񱳾�ͼƬ(&B)"
      End
      Begin VB.Menu MnuRightBar6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRightReLogin 
         Caption         =   "ע��(&R)"
      End
      Begin VB.Menu MnuRightExit 
         Caption         =   "�˳�(&X)"
      End
   End
End
Attribute VB_Name = "frmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurTime As Date                        '��ǰԤ����ʱ�����.
Private mblnFirst As Boolean
Private mlngMainMenu As Long                    '������Ĳ˵���ϵ�ľ��
Private mblnRemote As Boolean '�Ƿ���Զ��
'----����˵��----
'    ���޸���Ա��������
'1.���ܲ˵��Ĳ˵�ID��(90000001|10001)��ʼ
'2.���ڲ˵��Ĳ˵�ID��(99990001|30001)��ʼ(���ڲ˵��µ���Ϊ���ӵĲ˵�,������ʾ��ǰ�Ѵ�ģ��Ķ�̬�˵�)
'3.�������ܵĲ˵��Ĳ˵�ID��(99999901|65001)��ʼ
'4.ֻ�����һ���ָ����˵�,�����ڲ˵���(99999999|65535)
Private mblnHide As Boolean '�Ƿ���ʾ������
Private mclsAppTool As New zl9AppTool.clsAppTool
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1

Public Property Get frmHide() As Boolean
    frmHide = mblnHide
End Property

Public Property Get ObjLogin() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    Set ObjLogin = gobjRelogin
End Property

Public Property Get mobjEmr() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    Set mobjEmr = gobjRelogin.EMR
End Property

Private Sub cbrThis_Resize()
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim lngInstanceNo As Long
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    
    mnuOrderMenu.Checked = Val(zlDatabase.GetPara("zlMdiMenuArray")) <> 0
    mnuToolShowDisReport.Checked = IIf(Val(zlDatabase.GetPara("��ʾͣ�ñ���")) = 0, False, True)
    mnuRightShowDisReport.Checked = mnuToolShowDisReport.Checked
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
    
    '�˶α����ڴ���ͬ��ʺ�(����Ϣ֪ͨ����ZlAppTool����,ִ���亯��--GetUserInfoʱ����)
    MnuToolIndividuation.Checked = IIf(Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0, False, True)
    MnuToolNotify.Checked = IIf(Val(zlDatabase.GetPara("�����ʼ���Ϣ")) = 0, True, False)
    mnuToolTester.Checked = IIf(GetSetting("ZLSOFT", "����ȫ��", "SQLTest", 0) = 0, False, True)
    mnuRightTester.Checked = mnuToolTester.Checked
    mnuRightIndividuation.Checked = MnuToolIndividuation.Checked
    MnuToolNotify_Click
    mnuRightNotify.Checked = MnuToolNotify.Checked
    
    Me.stbThis.Panels(2).Text = ""
    stbThis.Panels(3).Text = IIf(gstrNodeName = "-", "", "Ժ����" & gstrNodeName)
    Me.stbThis.Panels(4).Text = gobjRelogin.DBUser & IIf(gobjRelogin.ServerName = "", "", "@" & gobjRelogin.ServerName) & IIf(zlDatabase.CheckRAC(lngInstanceNo), "(RAC:" & lngInstanceNo & ")", "")
    Me.stbThis.Panels(5).Text = gstrUserName
    Me.stbThis.Panels(6).Text = gstrDeptName
    Call SetMainForm(Me)                                '��ʼ������������������
    Call InitEvn
    Call LoadUsual
    Call LoadHistory
    
    '���˺�:�����ⲿ����
    '2007/08/22
    Call LoadOutTools
    
    Call Form_Resize
    
    '���ֻ��һ����ģ��,���
    On Error Resume Next
    With grsMenus
        .Filter = "ģ��<>0 And ����=0"
        If Not .EOF Then
            If .RecordCount = 1 Then
                Dim LngFind As Long, lngMenu As Long
                For LngFind = 0 To CollMenu.Count - 1
                    If CollMenu("K_" & LngFind)(Menu_Modul) = !ģ�� Then
                        lngMenu = CollMenu("K_" & LngFind)(Menu_ID)
                        Exit For
                    End If
                Next
                
                If lngMenu <> 0 Then Call MenuProc(Me.hwnd, WM_COMMAND, lngMenu, 0)
            End If
        End If
        .Filter = 0
    End With
    
    Call CheckWinVersion
    
    '������Ϣ����ƽ̨�ͻ����շ�����
    '------------------------------------------------------------------------------------------------------------------
    If ConnectMip(Me.hwnd) = True Then
        Set mclsMipModule = New zl9ComLib.clsMipModule
        Call mclsMipModule.InitMessage(0, 0, "")
        Call AddMipModule(mclsMipModule)
    End If
    
    '�����Զ����ѷ���
    mclsAppTool.CodeMan 0, 5, gcnOracle, Me, gstrDbUser
    If mblnHide Then Me.Hide '���ⲿ���ã�����������,by �¶�
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static StrPass As String                                '��������(Open zlReport.ReportMan )
    Dim LngFind As Long, BlnExist As Boolean, LngUpperMenu As Long, lngMenu As Long

    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        StrPass = ""
        Exit Sub
    End If

    If KeyCode <> vbKeyReturn Then
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then StrPass = StrPass & UCase(Chr(KeyCode))

        If StrPass = "OPEN ZLREPORT REPORTMAN" Then
            If OwnerUser(gstrDbUser) Then
                StrPass = ""
                
                If FindWindow(vbNullString, "�������") <> 0 Then Exit Sub
                If MsgBox("��ȷ��Ҫ�����Զ��屨������", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call ExecuteFunc(0, "ZL9REPORT", �˵���׼.�������ܲ˵�)
                SetParent FindWindow(vbNullString, "�������"), Me.hwnd
            End If
        End If
    End If
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim intGrant As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTitle As String, strTag As String
    mblnFirst = True
    
    On Error Resume Next
    strTitle = zlRegInfo("��Ʒ����")
    strTag = ""
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0) & IIf(strTag = "", "", "(" & strTag & ")")
    
    '�ж��Ƿ���Ȩ��ʹ����Ϣ�շ�����
    Me.Caption = gstrUserName & "-" & strTitle
    Call CheckTools
    RestoreWinState Me
    Call ApplyOEM_Picture(Me, "Icon")
    
    '��ȡ������Сֵ
    gLngMinH = Screen.Height - 400
    gLngMinW = Screen.Width
    gLngMaxH = gLngMinH
    gLngMaxW = gLngMinW
    
    Dim LngHdl As Long
'    'ȡϵͳ������
    Me.Width = gLngMinW
    Me.Height = gLngMinH
    
    Call InitEvn
    
    '���ϵͳ
    LngHdl = GetSubMenu(GetMenu(Me.hwnd), 0)
    Call InsertMenu(LngHdl, MF_BYPOSITION, MF_STRING, 99999999, "���Բ˵�(&T)")
    Me.Tag = GetMenuItemID(LngHdl, GetMenuItemCount(LngHdl) - 1)
    Call DeleteMenu(LngHdl, GetMenuItemCount(LngHdl) - 1, MF_BYPOSITION)
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    If Me.Tag <> 99999999 Then
        �˵���׼.���ܲ˵� = 10001
        �˵���׼.���ڲ˵� = 30001
        �˵���׼.�������ܲ˵� = 65001
        �˵���׼.�ָ��˵� = 65535
    Else
        �˵���׼.���ܲ˵� = 90000001
        �˵���׼.���ڲ˵� = 99990001
        �˵���׼.�������ܲ˵� = 99999901
        �˵���׼.�ָ��˵� = 99999999
    End If
    
    '�������ݿ����Ӹ���ӡ����
    IniPrintMode gcnOracle, gstrDbUser
    
    '�����жϻỰ���Ƿ�����Ϣ����������
    'select ����ֵ from zloptions where ������ =17
    strSQL = "select ����ֵ from zloptions where ������ =17"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж���ѯ�������Ƿ���")
    If rsTemp.RecordCount = 1 Then
        If NVL(rsTemp!����ֵ) <> "" Then
            '������ѯ������,�ر�TIME
            tmrUpdateConnect.Enabled = False
        Else
            'û����ѯ������,ʹ��TIME���� Ԥ�������
            tmrUpdateConnect.Enabled = True
            tmrUpdateConnect.Interval = 30000
            mCurTime = Now
        End If
    Else
        'û����ѯ������,ʹ��TIME���� Ԥ�������
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
    
    '�ⲿ���õĴ���,by �¶�
    mblnHide = False
    If gstrCommand <> "" Then Call DoCommand
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogInAfter
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� LogInAfter ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If

    '��ʼ������
    InitWinsock
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    
    With LvwList
        .Top = cbrThis.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
    With PicBackBitmap
        .Top = cbrThis.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnCloaseWin As Boolean
    
    blnCloaseWin = Val(zlDatabase.GetPara("�ر�Windows")) <> 0
    Set CollOpenWindowHdl = New Collection
    Set CollMenu = New Collection
    
    On Error Resume Next
    '�ָ�����ԭ�����ĵ�ַ
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    '�������ҽ�����Լ�ҵ����
    Call CloseChildWindows(Me)
    '������Ϣ����
    Call DisConnectMip
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Call gobjRelogin.Dispose '��Ҫ��ж�ض���
    Set gobjRelogin = Nothing
    SaveSetting "ZLSOFT", "����ȫ��", "SQLTest", 0
    '�������Ĳ���ֵ
    zlDatabase.ClearParaCache
    Call ShutDown(blnCloaseWin)
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 Then gcnOracle.Close
        Set gcnOracle = Nothing
    End If
    ReDim Preserve gobjCls(0)
    ReDim Preserve gstrObj(0)
End Sub

Private Sub HistoryItem_Click(Index As Integer)
    Dim strϵͳ As String, str��� As String
    strϵͳ = Split(HistoryItem(Index).Tag, ",")(0)
    str��� = Split(HistoryItem(Index).Tag, ",")(1)
    Debug.Print strϵͳ & ";" & str���
    With grsMenus
        .Filter = "ϵͳ=" & strϵͳ & " And ģ��=" & str���
        If .RecordCount <> 0 Then
            Call AddHistory(!ϵͳ & "," & !ģ��)
            Call LoadHistory
            .Filter = "ϵͳ=" & strϵͳ & " And ģ��=" & str���
            Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value)
        End If
        .Filter = 0
    End With
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuRightMenu, 2
End Sub

Public Sub MenuPrint(rsMenuList As ADODB.Recordset, intOutMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��ӡԤ��
    '������    �����ʽ
    '���أ�
    '---------------------------------------------------
    Dim objCol As Column, intCol As Integer
    Dim objPrint As New zlPrintDbGrd

    With dgdList
        For intCol = 1 To .Columns.Count - 1
            .Columns.Remove 0
        Next
        Set objCol = .Columns(0)
        objCol.Caption = "���"
        objCol.DataField = "ID"
        objCol.Alignment = dbgCenter
        objCol.Width = 500

        Set objCol = .Columns.Add(.Columns.Count)
        objCol.Caption = "����"
        objCol.DataField = "����"
        objCol.Alignment = dbgLeft
        objCol.Width = 1400

        Set objCol = .Columns.Add(.Columns.Count)
        objCol.Caption = "˵��"
        objCol.DataField = "˵��"
        objCol.Alignment = dbgLeft
        objCol.Width = 8000

        .HoldFields
    End With
    rsMenuList.Filter = 0
    Set dgdList.DataSource = rsMenuList

    '----------------------------------------------------

    If rsMenuList.EOF Or rsMenuList.BOF Then Exit Sub
    If InStr(1, Caption, "-") = 0 Then
        objPrint.Title.Text = Caption & "�����嵥"
    Else
        objPrint.Title.Text = Mid(Caption, 1, InStr(1, Caption, "-") - 1) & "�����嵥"
    End If

    Set objPrint.BodyGrid = dgdList
    Set objPrint.DataSource = rsMenuList

    If intOutMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewDBGrd objPrint, 1
        Case 2
            zlPrintOrViewDBGrd objPrint, 2
        Case 3
            zlPrintOrViewDBGrd objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrViewDBGrd objPrint, intOutMode
    End If

    Set dgdList.DataSource = Nothing
End Sub

Private Sub LoadMenuPortrait()
    Dim objNode As Node
    Dim LngMenuID As Long                           '�˵�ID��
    Dim LngLoop As Long                             'ѭ������
    Dim LngInsertMenu As Long                       '�����˵����
    Dim LngUpperMenu As Long                        '�ϼ��˵����
    Dim StrHotKey As String                         '��ݼ�
    On Error Resume Next
    '�������й��ܲ˵�
    '--�˵�ID��90000001��ʼ--
    
    LngMenuID = �˵���׼.���ܲ˵�
    Set CollMenu = New Collection                   '������Ӳ˵��������Ϣ
    TvwMenu.Nodes.Clear
    mlngMainMenu = GetMenu(Me.hwnd)
    mlngMainMenu = GetSubMenu(mlngMainMenu, 1)        '��ȡ�����Ӳ˵�
    
    With grsMenus
        Do While Not .EOF
            Err = 0
            If .Fields("�ϼ�") = 0 Then
                Set objNode = Me.TvwMenu.Nodes.Add(, , "_" & !���, !�̱���)
            Else
                Set objNode = Me.TvwMenu.Nodes.Add("_" & !�ϼ�, 4, "_" & !���, !�̱���)
            End If
            
            If Err = 0 Then
                    
                '�����ϼ��˵����
                LngUpperMenu = mlngMainMenu
                If Val(!�ϼ�) <> 0 Then
                    For LngLoop = 0 To CollMenu.Count - 1
                        If CollMenu("K_" & LngLoop)(Menu_Code) = !�ϼ� Then
                            LngUpperMenu = CollMenu("K_" & LngLoop)(Menu_Hdl)
                            Exit For
                        End If
                    Next
                End If
                
                StrHotKey = UCase(IIf(IsNull(!���), "", !���))
                StrHotKey = !�̱��� & IIf(StrHotKey = "", "", "(&" & StrHotKey & ")")
                '��Ӳ˵���(���ģ��ֵΪ��,��Ϊ�˵���;������ӵ����˵�)
                If !ģ�� = 0 Then
                    LngInsertMenu = CreatePopupMenu()
                    CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                Else
                    If !���� = 1 And Val(!�Ƿ�ͣ��) = 1 Then
                        If mnuToolShowDisReport.Checked Then
                            StrHotKey = StrHotKey & "(ͣ��)"
                            LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                            CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                            LngMenuID = LngMenuID + 1
                        End If
                    Else
                        LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                        CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                        LngMenuID = LngMenuID + 1
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    '�����е����˵���mnuOper�˵���,��Ϊ���¼��˵�
    Dim IntMenuLocate As Integer
    IntMenuLocate = 1
    For LngLoop = 0 To CollMenu.Count - 1
        If CollMenu("K_" & LngLoop)(Menu_Modul) = 0 Then
            StrHotKey = CollMenu("K_" & LngLoop)(Menu_Caption)              '�̱��⼰��ݼ�
            Call InsertMenu(CollMenu("K_" & LngLoop)(Menu_UpperHdl), IntMenuLocate, MF_BYPOSITION + MF_POPUP, CollMenu("K_" & LngLoop)(Menu_Hdl), StrHotKey)
            IntMenuLocate = IntMenuLocate + 1
        End If
    Next
    
    'ɾ��ȱʡ�˵�
    Call DeleteMenu(mlngMainMenu, 0, MF_BYPOSITION)
    
    'ˢ�²˵�
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)

    '���ô��庯���ĵ�ַ
    LngAddFunc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)
    '�ָ����˵����
    mlngMainMenu = GetMenu(Me.hwnd)
End Sub

Private Sub LoadMenuLandscape()
    Dim objNode As Node
    Dim LngMenuID As Long                           '�˵�ID��
    Dim LngLoop As Long                             'ѭ������
    Dim LngInsertMenu As Long                       '�����˵����
    Dim LngUpperMenu As Long                        '�ϼ��˵����
    Dim StrHotKey As String                         '��ݼ�
    On Error Resume Next
    '�������й��ܲ˵�
    '--�˵�ID��90000001��ʼ--
    
    LngMenuID = �˵���׼.���ܲ˵�
    Set CollMenu = New Collection                   '������Ӳ˵��������Ϣ
    TvwMenu.Nodes.Clear
    mlngMainMenu = GetMenu(Me.hwnd)
    'ɾ��"����"�˵�
    Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
    
    'ֱ�Ӷ����˵��������Ӳ���,��ʵ�ֺ������й��ܲ˵��ķ�ʽ
    With grsMenus
        Do While Not .EOF
            Err = 0
            If .Fields("�ϼ�") = 0 Then
                Set objNode = Me.TvwMenu.Nodes.Add(, , "_" & !���, !�̱���)
            Else
                Set objNode = Me.TvwMenu.Nodes.Add("_" & !�ϼ�, 4, "_" & !���, !�̱���)
            End If
            
            If Err = 0 Then
                    
                '�����ϼ��˵����
                LngUpperMenu = mlngMainMenu
                If Val(!�ϼ�) <> 0 Then
                    For LngLoop = 0 To CollMenu.Count - 1
                        If CollMenu("K_" & LngLoop)(Menu_Code) = !�ϼ� Then
                            LngUpperMenu = CollMenu("K_" & LngLoop)(Menu_Hdl)
                            Exit For
                        End If
                    Next
                End If
                
                StrHotKey = UCase(IIf(IsNull(!���), "", !���))
                StrHotKey = !�̱��� & IIf(StrHotKey = "", "", "(&" & StrHotKey & ")")
                '��Ӳ˵���(���ģ��ֵΪ��,��Ϊ�˵���;������ӵ����˵�)
                If !ģ�� = 0 Then
                    LngInsertMenu = CreatePopupMenu()
                    CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                Else
                    If !���� = 1 And Val(!�Ƿ�ͣ��) = 1 Then
                        If mnuToolShowDisReport.Checked Then
                            StrHotKey = StrHotKey & "(ͣ��)"
                            LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                            CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                            LngMenuID = LngMenuID + 1
                        End If
                    Else
                        LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                        CollMenu.Add Array(LngInsertMenu, .Fields("���").Value, .Fields("ģ��").Value, IIf(IsNull(!����), "", .Fields("����").Value), LngUpperMenu, StrHotKey, IIf(!ģ�� = 0, 0, LngMenuID), .Fields("ϵͳ").Value), "K_" & CollMenu.Count
                        LngMenuID = LngMenuID + 1
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    '�����е����˵���mnuOper�˵���,��Ϊ���¼��˵�
    Dim IntMenuLocate As Integer
    IntMenuLocate = 1
    For LngLoop = 0 To CollMenu.Count - 1
        If CollMenu("K_" & LngLoop)(Menu_Modul) = 0 Then
            StrHotKey = CollMenu("K_" & LngLoop)(Menu_Caption)              '�̱��⼰��ݼ�
            Call InsertMenu(CollMenu("K_" & LngLoop)(Menu_UpperHdl), IntMenuLocate, MF_BYPOSITION + MF_POPUP, CollMenu("K_" & LngLoop)(Menu_Hdl), StrHotKey)
            IntMenuLocate = IntMenuLocate + 1
        End If
    Next
    
    'ˢ�²˵�
    Call SetMenu(Me.hwnd, mlngMainMenu)
    Call DrawMenuBar(Me.hwnd)
    
    '���ô��庯���ĵ�ַ
    LngAddFunc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)
    '�ָ����˵����
    mlngMainMenu = GetMenu(Me.hwnd)
End Sub

Public Function Show����(ByVal ChildObj As Object)
    Dim LngWin As Long, BlnIn As Boolean, LngCount As Long
    Dim LngInsertMenu As Long, StrCaption As String
    Dim LngMenuCount As Integer, ClientRect As RECT, ClientPT As POINTAPI
    If grsMenus.State = 0 Then Exit Function
    If grsMenus.EOF Then Exit Function
    
    With grsMenus
        .MoveFirst
        .Find "����='" & ChildObj.Caption & "'"
        If .EOF Then
            .MoveFirst

            '������ڹ���
            If Trim(ChildObj.Caption) = "" Then Exit Function
            If InStr(1, "�Զ��屨�����,�ֵ������,��Ϣ�շ�����", ChildObj.Caption) <> 0 Then GoTo Normal
            Exit Function
        End If
    End With
    
Normal:                                                     '��������
    SetParent ChildObj.hwnd, Me.hwnd
    StrCaption = ChildObj.Caption
    '�ָ��Ӵ���ĸ߶�(��ȥ������ı������߶ȼ��˵��߶�)
    ClientPT.x = 0
    ClientPT.y = 0
    Call ClientToScreen(Me.hwnd, ClientPT)
    ChildObj.Top = ChildObj.Top - (ClientPT.y * 30)
    
    '�Ѵ��������뼯��
    BlnIn = False
    For LngWin = 0 To CollOpenWindowHdl.Count - 1
        If ChildObj.hwnd = CollOpenWindowHdl("K_" & LngWin)(0) Then
            BlnIn = True
            Exit For
        End If
    Next
    LngCount = CollOpenWindowHdl.Count
    
    If BlnIn = False Then
        CollOpenWindowHdl.Add Array(ChildObj.hwnd, ChildObj.Caption, �˵���׼.���ڲ˵� + CollOpenWindowHdl.Count), "K_" & LngCount
        
        grsMenus.Filter = "�ϼ� =0"
        LngMenuCount = IIf(mnuOrderMenu.Checked, grsMenus.RecordCount, 1) + IIf(History.Visible, 3, 2)
        grsMenus.Filter = 0
        
        LngInsertMenu = GetSubMenu(mlngMainMenu, LngMenuCount)           '��ȡ�����Ӳ˵�
        
        '����˵�
        If CollOpenWindowHdl.Count = 1 Then
            '����ָ��˵���
            Call InsertMenu(LngInsertMenu, MF_BYPOSITION, MF_SEPARATOR, �˵���׼.�ָ��˵�, "")
        End If
        '���봰�ڲ˵���
        Call InsertMenu(LngInsertMenu, MF_BYPOSITION, MF_STRING, �˵���׼.���ڲ˵� + CollOpenWindowHdl.Count - 1, StrCaption)
    End If
    
End Function

Public Sub Shut����(ByVal ObjFrm As Object)
    Dim LngDeleteMenu As Long, LngMenuCount As Integer
    Dim IntChange As Integer, IntDelete As Integer
    On Error Resume Next
        
    With grsMenus
        .Filter = "�ϼ� =0"
        LngMenuCount = IIf(mnuOrderMenu.Checked, .RecordCount, 1) + IIf(History.Visible, 3, 2)
        .Filter = 0
        
        .MoveFirst
        .Find "����='" & ObjFrm.Caption & "'"
        If .EOF Then
            .MoveFirst

            '������ڹ���
            If Trim(ObjFrm.Caption) = "" Then Exit Sub
            If InStr(1, "�Զ��屨�����,�ֵ������,��Ϣ�շ�����", ObjFrm.Caption) = 0 Then Exit Sub
        End If
    End With
    
    LngDeleteMenu = GetSubMenu(mlngMainMenu, LngMenuCount)           '��ȡ�����Ӳ˵�
    
    '--�������--
    For IntChange = 0 To CollOpenWindowHdl.Count - 1
        If CollOpenWindowHdl("K_" & IntChange)(1) = ObjFrm.Caption Then IntDelete = IntChange: Exit For
    Next
    
    If IntChange > CollOpenWindowHdl.Count Then Exit Sub
    '�����޸ĺ��
    For IntChange = IntChange To CollOpenWindowHdl.Count - 1
        CollOpenWindowHdl.Remove "K_" & IntChange
        CollOpenWindowHdl.Add CollOpenWindowHdl("K_" & IntChange + 1), "K_" & IntChange
    Next
    CollOpenWindowHdl.Remove "K_" & CollOpenWindowHdl.Count
    
    'ɾ����Ӧ�˵�
    Call DeleteMenu(LngDeleteMenu, IntDelete + 2, MF_BYPOSITION)
    
    '--�����Ӧ�˵�--
    If CollOpenWindowHdl.Count = 0 Then
        '����ָ��˵���
        Call DeleteMenu(LngDeleteMenu, 1, MF_BYPOSITION)
    End If
End Sub

Private Sub InitEvn()
    Dim StrPicPath As String, BlnShow As Boolean
    Dim LngColor As Long
    
    StrPicPath = zlDatabase.GetPara("zlMdiBackPic")
    
    If Trim(StrPicPath) <> "" Then
        '�û�ѡ��ͼƬ,�����Ƿ�����
        On Error Resume Next
        Err = 0
        BlnShow = False
        
        ImgTry.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "��ʾ����ͼƬʱ���������󣡣��ָ�ΪȱʡͼƬ��", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
        If BlnShow Then PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Else
        BlnShow = True
    End If
    
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '�ָ�ԭ�����õ�ͼƬ
    'ImgTry.Picture = LoadResPicture(101, 0) '�˵���ʶ
    'ȡ����ɫ
    LngColor = Val(zlDatabase.GetPara("zlMdiFontColor"))
    If LngColor <> -1 Then
        LvwList.ForeColor = LngColor
    End If
End Sub

Private Sub mclsMipModule_ConnectStateChanged(ByVal IsConnected As Boolean)
    '����״̬�Ѿ��仯
    If IsConnected Then
        tmrUpdateConnect.Enabled = False
    Else
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
End Sub

Private Sub mclsMipModule_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara)
End Sub

Private Sub mclsMipModule_OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara, True)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMessageConent As String)
    Select Case UCase(strMessageItemKey)
    '--------------------------------------------------------------------------------------------------------------
    Case "ZLHIS_PUB_005"            '��Ʒ����֪ͨ
        Call gobjRelogin.UpdateClient
    End Select

End Sub

Private Sub mnuFileExcel_Click()
    MenuPrint grsMenus, 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    MenuPrint grsMenus, 2
End Sub

Private Sub mnuFilePrint_Click()
    MenuPrint grsMenus, 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileReLogin_Click()
    If MsgBox("��ȷ��Ҫע����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ReLogin
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub

Private Sub mnuHelpWebForum_Click()
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub


Private Sub mnuOrderMenu_Click()
    Dim IntMenus As Integer, IntLastOrder As Integer, LngAddMenu As Long
    Dim FrmThis As Form, lngErr As Long, ClsClose As Object, intMenuCount As Integer
    
    grsMenus.Filter = 0
    
    IntLastOrder = IIf(mnuOrderMenu.Checked, 1, 0)
    mnuOrderMenu.Checked = Not mnuOrderMenu.Checked
    Call zlDatabase.SetPara("zlMdiMenuArray", IIf(mnuOrderMenu.Checked, 1, 0))
    If IntLastOrder = IIf(mnuOrderMenu.Checked, 1, 0) Then Exit Sub
    
    '�ָ�����ԭ�����ĵ�ַ
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    
    'ѭ��ɾ�����в˵�,������
    If IntLastOrder = 1 Then
        intMenuCount = GetMenuItemCount(mlngMainMenu)
        intMenuCount = intMenuCount - 6
        For IntMenus = 1 To intMenuCount
            Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
        Next
    End If
    
    '�����Ϊ�������з�ʽ,����������һ�������˵�
    If mnuOrderMenu.Checked = False Then
        LngAddMenu = CreatePopupMenu()
        Call InsertMenu(mlngMainMenu, 1, MF_STRING + MF_POPUP + MF_BYPOSITION, LngAddMenu, "����(&O)")
        mlngMainMenu = LngAddMenu
        LngAddMenu = CreateMenu()
        Call InsertMenu(mlngMainMenu, 0, MF_STRING, LngAddMenu, "Default")
    End If
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    '���Ӳ˵�
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
    Call LoadHistory
End Sub

Private Sub mnuRepairClientUpdate_Click()
    If MsgBox("�����������¼�Ȿ�������������Ա����������������޸������޸�������в�����������ע�ᡣ��ȷ��Ҫ���пͻ����޸���", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call gobjRelogin.UpdateClient(True)
    End If
End Sub

Private Sub mnuRepairComponent_Click()
    '--���ע���[��������]--
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
    MsgBox "���������ϣ����иĶ������µ�¼����Ч��", vbInformation, gstrSysName
End Sub

Private Sub mnuRepairIndividuationClear_Click()
    Dim strSQL As String, rsTmp As Recordset
    Dim strAnalyseComputer As String
    
    If MsgBox("�����������ZLHIS��ص�ע���������Լ����ݿ��д洢�ı��ˡ�������������Ʒ��ع��ܽ�������ȱʡֵ���У���ȷ��Ҫ������", vbYesNo + vbDefaultButton2 + vbQuestion, "������������쳣") = vbYes Then
        strSQL = "Select Distinct ���� From zlPrograms Where ���� Is Not Null"
        On Error GoTo ErrHand
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������������쳣")
        Do While Not rsTmp.EOF
            Call DelWinState(Me, rsTmp!���� & "")
            rsTmp.MoveNext
        Loop
        strAnalyseComputer = OS.ComputerName
        strSQL = "Zl_zluserparas_Clear('" & gstrDbUser & "','" & strAnalyseComputer & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
        MsgBox "����ɹ�����رճ������½��룬ȷ���Ƿ��������쳣���⡣", vbInformation, "������������쳣"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub MnuRightAbout_Click()
    mnuHelpAbout_Click
End Sub

Private Sub mnuRightClientUpdate_Click()
    Call mnuRepairClientUpdate_Click
End Sub

Private Sub MnuRightComponent_Click()
    mnuRepairComponent_Click
End Sub

Private Sub mnuRightDictonary_Click()
    mnuToolDictonary_Click
End Sub

Private Sub MnuRightExcel_Click()
    Call mnuToolExcel_Click
End Sub

Private Sub MnuRightExit_Click()
    mnuFileExit_Click
End Sub

Private Sub MnuRightHistory_Click()
    Call mnuToolHistory_Click
End Sub

Private Sub mnuRightIndividuation_Click()
    MnuToolIndividuation_Click
End Sub

Private Sub mnuRightIndividuationClear_Click()
    mnuRepairIndividuationClear_Click
End Sub

Private Sub mnuRightMessage_Click()
    mnuToolMessage_Click
End Sub

Private Sub mnuRightNotice_Click()
    Call mnuToolNotice_Click
End Sub

Private Sub mnuRightNotify_Click()
    MnuToolNotify_Click
End Sub

Private Sub MnuRightReLogin_Click()
    mnuFileReLogin_Click
End Sub

Private Sub mnuRightShowDisReport_Click()
    Call mnuToolShowDisReport_Click
End Sub

Private Sub MnuRightStyle_Click()
    mnuTooleSelect_Click
End Sub

Private Sub mnuRightTester_Click()
    mnuToolTester_Click
End Sub

Private Sub mnurightBackBmp_Click()
    Dim BlnShow As Boolean              '�ܷ�������ʾ
    Dim StrPicPath As String            '����ͼƬ·��
    '--���û�ѡ�񱳾�ͼƬ--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .Filter = "����ͼƬ (*.bmp;*.jpg)|*.bmp;*.jpg"
        .ShowOpen
        
        '�û�ѡ��ͼƬ,�����Ƿ�����
        On Error Resume Next
        Err = 0
        BlnShow = False
        
        StrPicPath = .FileName
        ImgTry.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "����ѡ���ͼƬ�ļ���������ʾ��", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
    End With
    
    PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '����ͼƬλ�ù��´���ȡ
    Call zlDatabase.SetPara("zlMdiBackPic", StrPicPath)
    '�ָ�ԭ�����õ�ͼƬ
    'ImgTry.Picture = LoadResPicture(101, 0) '�˵���ʶ
ErrHand:
End Sub

Private Sub mnurightSetColor_Click()
    '--���û�ѡ��������ɫ--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .ShowColor
        
        LvwList.ForeColor = .Color
        
        '��������ɫ���´���ȡ
        Call zlDatabase.SetPara("zlMdiFontColor", .Color)
    End With
ErrHand:
End Sub

Private Sub mnuToolDictonary_Click()
    mclsAppTool.CodeMan 0, 1, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolExcel_Click()
    Dim ObjExcel As Object, strHaveSys As String
    
    If gstrUserName = "" Then
        MsgBox "��Ϊ����Ա���ö�Ӧ���û�����ʹ�ñ����ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strHaveSys = gobjRelogin.Systems
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Zl9Excel.ClsExcel")
    If Err <> 0 Then
        MsgBox "�޷�����EXCEL��������������ʹ��EXCEL����", vbInformation, gstrSysName
        Exit Sub
    End If
    Call ObjExcel.CodeMan(0, 0, gcnOracle, Me, gstrDbUser)
    Call ObjExcel.SetHaveSys(strHaveSys)
    Call ObjExcel.ExcelReportMain
    Set ObjExcel = Nothing
End Sub

Private Sub mnuToolHistory_Click()
    Call zlDatabase.SetPara("���ʹ��ģ��", "")
    Call LoadHistory
End Sub

Private Sub MnuToolIndividuation_Click()
    MnuToolIndividuation.Checked = MnuToolIndividuation.Checked Xor True
    mnuRightIndividuation.Checked = MnuToolIndividuation.Checked
    Call zlDatabase.SetPara("ʹ�ø��Ի����", IIf(MnuToolIndividuation.Checked, "1", "0"))
    SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser, "ʹ�ø��Ի����", IIf(MnuToolIndividuation.Checked, "1", "0")
End Sub

Private Sub MnuToolIndividuationClear_Click()
    Dim strSQL As String, rsTmp As Recordset
    Dim strAnalyseComputer As String
    
    If MsgBox("�����������ZLHIS��ص�ע���������Լ����ݿ��д洢�ı��ˡ�������������Ʒ��ع��ܽ�������ȱʡֵ���У���ȷ��Ҫ������", vbYesNo + vbDefaultButton2 + vbQuestion, "������������쳣") = vbYes Then
        strSQL = "Select Distinct ���� From zlPrograms Where ���� Is Not Null"
        On Error GoTo ErrHand
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������������쳣")
        Do While Not rsTmp.EOF
            Call DelWinState(Me, rsTmp!���� & "")
            rsTmp.MoveNext
        Loop
        strAnalyseComputer = OS.ComputerName
        strSQL = "Zl_zluserparas_Clear('" & gstrDbUser & "','" & strAnalyseComputer & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
        MsgBox "����ɹ�����رճ������½��룬ȷ���Ƿ��������쳣���⡣", vbInformation, "������������쳣"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuToolMessage_Click()
    mclsAppTool.CodeMan 0, 2, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuTooleSelect_Click()
    mclsAppTool.CodeMan 0, 3, gcnOracle, Me, gstrDbUser, gstrMenuSys
    If Val(zlDatabase.GetPara("����Զ�̿���")) <> winSock.LocalPort Then
        Call InitWinsock
    End If
    If mclsAppTool.IsRestart Then
        mclsAppTool.IsRestart = False
        Call ReLogin
    Else
        Call ShutUsual
        Call LoadUsual
    End If
End Sub

Private Sub mnuToolNotice_Click()
    mclsAppTool.CodeMan 0, 6, gcnOracle, Me, gstrDbUser
End Sub

Private Sub MnuToolNotify_Click()
    MnuToolNotify.Checked = Not MnuToolNotify.Checked
    mnuRightNotify.Checked = MnuToolNotify.Checked
    Call zlDatabase.SetPara("�����ʼ���Ϣ", IIf(MnuToolNotify.Checked, "1", "0"))
    mclsAppTool.CodeMan 0, 4, gcnOracle, Me, gstrDbUser, IIf(MnuToolNotify.Checked = True, "Open", "Close")
End Sub

Private Sub mnuToolShowDisReport_Click()
    Dim IntMenus As Integer, intMenuCount As Integer
    Dim LngAddMenu As Long
    
    mnuToolShowDisReport.Checked = Not mnuToolShowDisReport.Checked
    mnuRightShowDisReport.Checked = mnuToolShowDisReport.Checked
    Call zlDatabase.SetPara("��ʾͣ�ñ���", IIf(mnuToolShowDisReport.Checked, 1, 0))

    grsMenus.Filter = 0
    
    '�ָ�����ԭ�����ĵ�ַ
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    
    'ѭ��ɾ�����в˵�,������
    intMenuCount = GetMenuItemCount(mlngMainMenu)
    If mnuOrderMenu.Checked Then
        intMenuCount = intMenuCount - 7
    Else
        intMenuCount = intMenuCount - 6
    End If
    For IntMenus = 1 To intMenuCount
        Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
    Next
    
    '�����Ϊ�������з�ʽ,����������һ�������˵�
    If Not mnuOrderMenu.Checked Then
        LngAddMenu = CreatePopupMenu()
        Call InsertMenu(mlngMainMenu, 1, MF_STRING + MF_POPUP + MF_BYPOSITION, LngAddMenu, "����(&O)")
        mlngMainMenu = LngAddMenu
        LngAddMenu = CreateMenu()
        Call InsertMenu(mlngMainMenu, 0, MF_STRING, LngAddMenu, "Default")
    End If
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    '���Ӳ˵�
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
End Sub
 
Private Sub mnuToolTester_Click()
    mnuToolTester.Checked = mnuToolTester.Checked Xor True
    mnuRightTester.Checked = mnuToolTester.Checked
    SaveSetting "ZLSOFT", "����ȫ��", "SQLTest", IIf(mnuToolTester.Checked, 1, 0)
End Sub

Private Sub mnuWindowList_Click()
    Dim RectThis As RECT
    
    Call GetClientRect(Me.hwnd, RectThis)
    Call CascadeWindows(Me.hwnd, 0, RectThis, 0, 0)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Dictionary"
        mnuToolDictonary_Click
    Case "Message"
        mnuToolMessage_Click
    Case "Choose"
        mnuTooleSelect_Click
    Case "Check"
        mnuRepairComponent_Click
    Case "����"
        '���˺�:2007/08/22
        '����:�����ⲿ����
        Call mnuToolOutToolSet_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Key = "����" Then
        '���˺�:2007/08/22
        '����:�����ⲿ����
        Call ExeCuteToolFile(ButtonMenu.Tag)
        Exit Sub
    End If
End Sub

Private Sub TbrUsual_ButtonClick(ByVal Button As MSComctlLib.Button)
    With grsMenus
        .Filter = "ϵͳ=" & Split(Button.Tag, ",")(0) & " And ģ��=" & Split(Button.Tag, ",")(1)
        If .RecordCount <> 0 Then
            Call AddHistory(!ϵͳ & "," & !ģ��)
            Call LoadHistory
            .Filter = "ϵͳ=" & Split(Button.Tag, ",")(0) & " And ģ��=" & Split(Button.Tag, ",")(1)
            Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value)
        End If
        .Filter = 0
    End With
End Sub

Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)  '�����̬����
End Sub

Public Sub LoadHistory()
    Dim strϵͳ As String, str��� As String
    Dim arrϵͳ As Variant, arr��� As Variant
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer
    Dim strValue As String
    
    '����ʷ��¼װ��˵�
    Call ClearHistoryMenu
    strValue = zlDatabase.GetPara("���ʹ��ģ��")
    If UBound(Split(strValue, "|")) < 1 Then Exit Sub
    strϵͳ = Trim(Split(strValue, "|")(0))
    str��� = Trim(Split(strValue, "|")(1))
    If strϵͳ = "" Or str��� = "" Then Exit Sub
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    If intϵͳ_Max > 8 Then intϵͳ_Max = 8 '���˸���ʷ��¼
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        
        With grsMenus
            .Filter = "ϵͳ=" & IIf(arrϵͳ(intϵͳ_Cur) = "", 0, arrϵͳ(intϵͳ_Cur)) & " And ģ��=" & arr���(int���_Cur)
            If .RecordCount <> 0 Then
                '����ȱʡֵ
                Load HistoryItem(HistoryItem.Count)
                With HistoryItem(HistoryItem.Count - 1)
                    .Caption = grsMenus!����
                    .Visible = True
                    .Enabled = True
                    .Tag = grsMenus!ϵͳ & "," & grsMenus!ģ��
                End With
            End If
            .Filter = 0
        End With
    Next
    If HistoryItem.UBound > 0 Then
        HistoryItem(0).Visible = False
    End If
End Sub

Private Sub ClearHistoryMenu()
    Dim MenuItem As Menu
    On Error Resume Next
    
    'ɾ����ʷ��¼�˵�
    For Each MenuItem In HistoryItem
        If MenuItem.Index <> 0 Then
            Unload MenuItem
        Else
            MenuItem.Visible = True
        End If
    Next
End Sub

Private Sub LoadUsual()
    Dim strϵͳ As String, str��� As String, strͼ�� As String, str���� As String
    Dim arrϵͳ As Variant, arr��� As Variant, arrͼ�� As Variant, arr���� As Variant
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer, intͼ��_Cur As Integer, int����_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer, intͼ��_Max As Integer, int����_Max As Integer
    Dim objButton As Button, strValue As String
    
    '���ӳ��ù���
    strValue = zlDatabase.GetPara("���ù���ģ��")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    strϵͳ = Trim(Split(strValue, "|")(0))
    str��� = Trim(Split(strValue, "|")(1))
    strͼ�� = Trim(Split(strValue, "|")(2))
    str���� = Trim(Split(strValue, "|")(3))
    If strϵͳ = "" Or str��� = "" Then Exit Sub
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    arrͼ�� = Split(strͼ��, ",")
    arr���� = Split(str����, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    intͼ��_Max = UBound(arrͼ��)
    int����_Max = UBound(arr����)
    
    '����ͼ��
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        intͼ��_Cur = intϵͳ_Cur
        int����_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        
        ImgUsualBlack.ImageHeight = 32
        ImgUsualBlack.ImageWidth = 32
        With grsMenus
            .Filter = "ϵͳ=" & arrϵͳ(intϵͳ_Cur) & " And ģ��=" & arr���(int���_Cur)
            If .RecordCount <> 0 Then
                '����ȱʡֵ
                If intͼ��_Cur <= intͼ��_Max Then
                    strͼ�� = arrͼ��(intͼ��_Cur)
                Else
                    strͼ�� = !ͼ��
                End If
                ImgUsualBlack.ListImages.Add , "K" & intϵͳ_Cur, GetPicDisp(strͼ��)
            End If
            .Filter = 0
        End With
    Next
    
    '���Ӱ�ť
    If ImgUsualBlack.ListImages.Count = 0 Then Exit Sub
    TbrUsual.Buttons.Clear
    Set TbrUsual.ImageList = ImgUsualBlack
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        int����_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        
        With grsMenus
            .Filter = "ϵͳ=" & arrϵͳ(intϵͳ_Cur) & " And ģ��=" & arr���(int���_Cur)
            If .RecordCount <> 0 Then
                '����ȱʡֵ
                strϵͳ = !ϵͳ
                str��� = !ģ��
                If int����_Cur <= int����_Max Then
                    str���� = arr����(int����_Cur)
                Else
                    str���� = !����
                End If
                Set objButton = TbrUsual.Buttons.Add()
                objButton.Caption = ""
                objButton.ToolTipText = str����
                objButton.Tag = strϵͳ & "," & str���
                objButton.Image = "K" & intϵͳ_Cur
                objButton.Key = "K" & intϵͳ_Cur
                objButton.Visible = True
            End If
            .Filter = 0
        End With
    Next
    DoEvents
    cbrThis.Bands(2).MinHeight = TbrUsual.Height
    Set cbrThis.Bands(2).Child = TbrUsual
    cbrThis.Bands(2).Visible = True
    DoEvents
End Sub

Private Sub ShutUsual()
    Dim intButton As Integer
    'ɾ�����г��ù���
    
    Set TbrUsual.ImageList = Nothing
    For intButton = 1 To TbrUsual.Buttons.Count
        TbrUsual.Buttons.Remove (1)
    Next
    ImgUsualBlack.ListImages.Clear
    cbrThis.Bands(2).Visible = False
    Call Form_Resize
End Sub

Private Sub CheckTools()
    Dim blnSplit As Boolean         '�Ƿ���ʾ�ָ���
    '��Ϣ�շ���EXCEL�����Ȩ�޿��ƣ�
    '1�������Ȩ���к��д˹���
    '2��������û�ӵ�д�Ȩ��
    '3����ʾ����������
    '��������ģ����жϸ��û��Ƿ�ӵ�д�Ȩ��
    
    '���߶�Ӧ˵��
    '��ӡ��Ԥ�������EXCEL  ,10,'���������嵥','����'
    'mnuToolDictonary       ,11,'�ֵ������','����'
    'mnuToolMessage         ,12,'��Ϣ�շ�����','����,������Ϣ'
    'mnuTooleSelect         ,13,'ϵͳѡ������','����'
    'mnuToolExcel           ,14,'EXCEL������','����,������ɾ,�������,����ϵͳ'
    'mnuToolUp              ,15,'���ز����ϴ�' ,'����'
    
    Dim intGrant As Integer
    
    '���������嵥
    mnuFilePrint.Visible = False
    mnuFilePreview.Visible = False
    mnuFileExcel.Visible = False
    tbrThis.Buttons("Print").Visible = False
    tbrThis.Buttons("Preview").Visible = False
    tbrThis.Buttons("printbar").Visible = False
    'Excel������
    mnuToolExcel.Visible = False
    MnuRightExcel.Visible = False
    '��Ϣ�շ�����
    mnuToolMessage.Visible = False
    MnuToolNotify.Visible = False
    mnuRightMessage.Visible = False
    mnuRightNotify.Visible = False
    tbrThis.Buttons("Message").Visible = False
    'ϵͳѡ������
    mnuTooleSelect.Visible = False
    MnuRightStyle.Visible = False
    tbrThis.Buttons("Choose").Visible = False
    '�ֵ������
    mnuToolDictonary.Visible = False
    mnuRightDictonary.Visible = False
    tbrThis.Buttons("Dictionary").Visible = False
    '��Ȼ,�ָ���һ����Ҫ��ֹ��,ֻҪ��������һ�����ܣ��ֵ������Ϣ�շ���EXCEL�����ϵͳѡ�������Ҫ��ʾ�ָ���
    blnSplit = False
    
    intGrant = zlRegTool '(GetUnitInfo("ע����"))
    If ((intGrant And 4) = 4) Then
        If InStr(1, GetPrivFunc(0, �����嵥.��Ϣ�շ�����), "����") <> 0 Then
            mnuToolMessage.Visible = True
            MnuToolNotify.Visible = True
            mnuRightMessage.Visible = True
            mnuRightNotify.Visible = True
            tbrThis.Buttons("Message").Visible = True
            blnSplit = True
        Else
            Call zlDatabase.SetPara("�����ʼ���Ϣ", "0")
        End If
    End If
    If ((intGrant And 8) = 8) Then
        If InStr(1, GetPrivFunc(0, �����嵥.EXCEL������), "����") Then
            mnuToolExcel.Visible = True
            MnuRightExcel.Visible = True
            blnSplit = True
        End If
    End If

    If InStr(1, GetPrivFunc(0, �����嵥.���������嵥), "����") Then
        mnuFilePrint.Visible = True
        mnuFilePreview.Visible = True
        mnuFileExcel.Visible = True
        tbrThis.Buttons("Print").Visible = True
        tbrThis.Buttons("Preview").Visible = True
        tbrThis.Buttons("printbar").Visible = True
    End If
    If InStr(1, GetPrivFunc(0, �����嵥.ϵͳѡ������), "����") Then
        mnuTooleSelect.Visible = True
        MnuRightStyle.Visible = True
        tbrThis.Buttons("Choose").Visible = True
        blnSplit = True
    End If
    If InStr(1, GetPrivFunc(0, �����嵥.�ֵ������), "����") Then
        mnuToolDictonary.Visible = True
        mnuRightDictonary.Visible = True
        tbrThis.Buttons("Dictonary").Visible = True
        blnSplit = True
    End If
    mnuToolSplit2.Visible = blnSplit
    MnuRightBar3.Visible = blnSplit
    
    '���û��"��Ϣ�շ���ϵͳѡ��ֵ����"
    tbrThis.Buttons("bar").Visible = (mnuToolDictonary.Visible Or mnuTooleSelect.Visible Or mnuToolMessage.Visible)
End Sub

Public Sub RunModual(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strPara As String, Optional ByVal blnReport As Boolean)
    '------------------------------------------------------------------------------------------------------
    '����:����ִ�б���,�˹�����Ϊ�Զ����ѵ��ö�д,by �¸���
    '����:lngSys ϵͳ���;lngModual ģ���
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHand
    
    With grsMenus
        If blnReport Then
            .Filter = "ϵͳ=" & lngSys & " AND ģ��=" & lngModual & " And ����=1"
        Else
            .Filter = "ϵͳ=" & lngSys & " AND ģ��=" & lngModual
        End If
        If .RecordCount = 0 Then .Filter = 0: Exit Sub
        If .Fields("ģ��").Value <> 0 Then
            Call ExecuteFunc(.Fields("ϵͳ").Value, IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value), .Fields("ģ��").Value, strPara)
        End If
        .Filter = 0
    End With
    
ErrHand:
    
End Sub


Private Sub mnuToolOutToolExecute_Click(Index As Integer)
    '���˺�:2007/08/22
    '���Ӷ��ⲿ���ߵ�ִ��
    Call ExeCuteToolFile(mnuToolOutToolExecute(Index).Tag)
End Sub
Private Sub ExeCuteToolFile(ByVal strFile As String)
    '-----------------------------------------------------------------------------------
    '����:ִ�й����ļ�
    '����:strFile-�ļ���
    '����:���˺�
    '����:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Err = 0: On Error GoTo ErrHand:
    If objFile.FileExists(strFile) = False Then
        MsgBox "�����ļ�:" & strFile & vbCrLf & "������,�����ѱ�ɾ��,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    Shell strFile, vbNormalFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub mnuToolOutToolSet_Click()
    Dim blnApply As Boolean
    '���˺�:2007/08/22
    '�����ⲿ���ߵ�����
    Call frm��������.ShowEdit(Me, blnApply)
    If blnApply = False Then Exit Sub
    Call LoadOutTools
End Sub
Private Function LoadOutTools() As Boolean
    '-----------------------------------------------------------------------------------
    '����:�����ⲿ����
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim i As Long
    Dim strReg As String, arrTemp As Variant, ArrTool As Variant
    Dim objButton As ButtonMenu
    Err = 0: On Error Resume Next
    '������ⲿ���߲˵�
    For i = 1 To mnuToolOutToolExecute.UBound
        Unload mnuToolOutToolExecute(i)
    Next
    
    '�����������
    Do While True
        If tbrThis.Buttons("����").ButtonMenus.Count = 0 Then Exit Do
        tbrThis.Buttons("����").ButtonMenus.Remove tbrThis.Buttons("����").ButtonMenus.Count
    Loop
    tbrThis.Buttons("����").Style = tbrDefault
    mnuToolOutToolExecute(0).Visible = False
    '���ع��߲˵�
    strReg = GetSetting("ZLSOFT", "����ȫ��\TOOLS", "TOOLFILES", "")
    If strReg = "" Then Exit Function
    ArrTool = Split(strReg, "|")
    For i = 0 To UBound(ArrTool)
        arrTemp = Split(ArrTool(i) & ",", ",")
        If arrTemp(0) <> "" And arrTemp(1) <> "" Then
            If i = 0 Then
                With mnuToolOutToolExecute(0)
                    .Caption = arrTemp(0) & "(&1)"
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            Else
                Load mnuToolOutToolExecute(i)
                With mnuToolOutToolExecute(i)
                    .Caption = arrTemp(0) & IIf(i + 1 > 9, "", "(&" & i + 1 & ")")
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            End If
            With tbrThis.Buttons("����").ButtonMenus
                Set objButton = .Add(, "K" & i, arrTemp(0))
                objButton.Tag = arrTemp(1)
            End With
            tbrThis.Buttons("����").Style = tbrDropdown
        End If
    Next
    LoadOutTools = True
End Function

Public Function GetCommand() As String
    '����:����ҵ�񲿼���ȡ�����в���,by �¶�
    '����:��
    GetCommand = gstrCommand
End Function

Private Sub DoCommand()
    '���ܣ��ⲿ���õ���̨ʱ�����ݴ����������ҵ�񲿼���,by �¶�
    '��������
    Dim i As Integer, lngModual As Long
    Dim varCmd As Variant
    On Error GoTo errH
    varCmd = Split(gstrCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "PROGRAM=*" Then
            lngModual = Val(Split(varCmd(i), "=")(1))
            grsMenus.Filter = "ģ��=" & lngModual
            If Not grsMenus.EOF Then
                Call RunModual(grsMenus!ϵͳ, lngModual, "")
                mblnHide = True
            End If
            grsMenus.Filter = 0
        End If
    Next
    Exit Sub
errH:
    
End Sub

Public Sub UnloadForm()
    '���ܣ��ⲿ���õ���̨����ҵ�񲿼���ҵ�񲿼����˳�ʱ��Ҫ���ô˺����رյ���̨��by �¶�
    '��������
    Unload Me
End Sub

Private Sub tmrUpdateConnect_Timer()
    'Ԥ��������
    If DateAdd("n", -30, Now) >= mCurTime Then '30���Ӽ��һ��
        tmrUpdateConnect.Enabled = False
        Call gobjRelogin.UpdateClient
        mCurTime = Now
        tmrUpdateConnect.Enabled = True
    End If
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object) As Boolean
     '����:�ر������Ӵ���
    Dim FrmThis     As Form, ClsClose As Object, IntCount As Integer, lngErr As Long
    Dim objInsure   As Object
    Dim blnOK       As Boolean
    
    On Error Resume Next
    blnOK = True
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                blnOK = False
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogOutBefore
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            blnOK = False
            MsgBox "zlPlugIn ��Ҳ���ִ�� LogOutBefore ʱ����" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    On Error Resume Next
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption Then Unload FrmThis
    Next
    '�ر����в����Ĵ���
    If Err.Number <> 0 Then Err.Clear
    lngErr = UBound(gstrObj)
    If Err.Number = 0 Then
        For IntCount = 0 To lngErr
            Set ClsClose = gobjCls(IntCount)
            blnOK = blnOK And ClsClose.CloseWindows
            Set gobjCls(IntCount) = Nothing
        Next
    End If
    '�ر�Ӧ�ù��߰������Ĵ���
    blnOK = blnOK And mclsAppTool.CloseWindows
    '�رչ��������Ĵ���
    blnOK = blnOK And CloseWindows
    Set objInsure = GetObject("", "zl9Insure.clsInsure")
    Call objInsure.Releaseme
    If Err.Number <> 0 Then Err.Clear
    CloseChildWindows = blnOK
End Function

Public Function GetPicDisp(Optional ByVal intIcon As Long = 0, Optional ByVal Blnģ�� As Boolean = True) As IPictureDisp
    '������:����
    '��������:2000-12-12
    '�õ�ͼƬ����

    On Error Resume Next
    If intIcon = 0 Then intIcon = IIf(Blnģ��, -5, -4)
    Select Case intIcon
    Case -1
        Set GetPicDisp = LoadResPicture("HELP", 1)
    Case -2
        Set GetPicDisp = LoadResPicture("RELOGIN", 1)
    Case -3
        Set GetPicDisp = LoadResPicture("EXIT", 1)
    Case -4
        Set GetPicDisp = LoadResPicture("DIRECTORY", 1)
    Case -5
        Set GetPicDisp = LoadResPicture("MODUL", 1)
    Case Else
        Set GetPicDisp = mclsAppTool.GetIcon(intIcon)
    End Select
End Function

Private Sub InitWinsock()
'����:��ȡ����,��ʼ��������
    Dim lngPort As Long
            
    On Error Resume Next
    
    lngPort = Val(zlDatabase.GetPara("����Զ�̿���"))
    mblnRemote = Not lngPort = -1
    winSock.Tag = "1"
    With winSock
        If mblnRemote Then
            .LocalPort = IIf(Val(lngPort) = 0, "1001", Val(lngPort))
            .Listen
        Else
            If .State <> sckClosed Then .Close
        End If
    End With
    winSock.Tag = ""
End Sub

Private Sub winSock_Close()
    If winSock.Tag = "" Then
        If winSock.State <> sckClosed And mblnRemote Then winSock.Close: winSock.Listen  '���¼���
    End If
End Sub

Private Sub winSock_ConnectionRequest(ByVal requestID As Long)
    If winSock.State <> sckClosed Then winSock.Close
    winSock.Accept requestID
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strMsg  As String
    
    winSock.GetData strData
    
    On Error GoTo errH
    If strData = "����Զ��" Then
                RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
                winSock.SendData "YES"
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    winSock.Close: winSock.Listen
    If winSock.Tag = "" Then
        Select Case Number
            Case 10053
                MsgBox "���ڳ�ʱ��û�в����������Զ��жϡ�", vbInformation, gstrSysName
            Case Else
                MsgBox Number & Description, vbInformation, gstrSysName
         End Select
    Else
        winSock.Tag = ""
    End If
End Sub


