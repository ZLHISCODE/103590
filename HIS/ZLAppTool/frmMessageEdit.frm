VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMessageEdit 
   Caption         =   "��Ϣ"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9510
   Icon            =   "frmMessageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Tag             =   "�ɱ仯��"
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
      Caption         =   "�ռ���(&R)"
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
               Caption         =   "��"
               Key             =   "Reply"
               Object.ToolTipText     =   "��"
               Object.Tag             =   "��"
               ImageKey        =   "Reply"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ReplyAll"
               Object.ToolTipText     =   "ȫ����"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ReplyAll"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ת��"
               Key             =   "Forward"
               Object.ToolTipText     =   "ת��"
               Object.Tag             =   "ת��"
               ImageKey        =   "Forward"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
               Object.ToolTipText     =   "����"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "ճ��"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "splite1"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Font"
               Object.ToolTipText     =   "����"
               Style           =   4
               Object.Width           =   1600
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Size"
               Object.ToolTipText     =   "�ֺ�"
               Style           =   4
               Object.Width           =   1100
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Bold"
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "б��"
               ImageKey        =   "Italic"
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnderLine"
               Object.ToolTipText     =   "�»���"
               ImageKey        =   "UnderLine"
               Style           =   1
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fore"
               Object.ToolTipText     =   "����ɫ"
               ImageKey        =   "Fore"
               Style           =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "����ɫ"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Object.ToolTipText     =   "�������"
               ImageKey        =   "Left"
               Style           =   2
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "���ж���"
               ImageKey        =   "Center"
               Style           =   2
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "���Ҷ���"
               ImageKey        =   "Right"
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Sig"
               Object.ToolTipText     =   "��Ŀ����"
               ImageKey        =   "Sig"
               Style           =   1
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Decrease"
               Object.ToolTipText     =   "����������"
               ImageKey        =   "Decrease"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Increase"
               Object.ToolTipText     =   "����������"
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
               Caption         =   "����"
               Key             =   "Send"
               Object.ToolTipText     =   "������Ϣ"
               Object.Tag             =   "����"
               ImageKey        =   "Send"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Save"
               Object.ToolTipText     =   "������Ϣ"
               Object.Tag             =   "����"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ҫ"
               Key             =   "High"
               Object.ToolTipText     =   "��Ҫ�ԣ���"
               Object.Tag             =   "��Ҫ"
               ImageKey        =   "High"
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ҫ"
               Key             =   "Low"
               Object.ToolTipText     =   "��Ҫ�ԣ���"
               Object.Tag             =   "��Ҫ"
               ImageKey        =   "Low"
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
      Caption         =   "����(&S)��"
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   2250
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ռ��ˣ�"
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   12
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ�䣺"
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����ˣ�"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   1530
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSend 
         Caption         =   "����(&E)"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "����(&T)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "ճ��(&P)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "���(&L)"
      End
      Begin VB.Menu mnuEditSel 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "����(&F)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "������һ��(&N)"
         Shortcut        =   +{F3}
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
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "��ʽ(&F)"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuFormatFore 
         Caption         =   "����ɫ(&O)"
      End
      Begin VB.Menu mnuFormatBack 
         Caption         =   "����ɫ(&B)"
      End
      Begin VB.Menu mnuFormatSig 
         Caption         =   "��Ŀ����(&S)"
      End
      Begin VB.Menu mnuFormatSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatLeft 
         Caption         =   "�������(&L)"
      End
      Begin VB.Menu mnuFormatCenter 
         Caption         =   "���ж���(&C)"
      End
      Begin VB.Menu mnuFormatRight 
         Caption         =   "���Ҷ���(&R)"
      End
      Begin VB.Menu mnuFormatSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatDecrease 
         Caption         =   "����������(&D)"
      End
      Begin VB.Menu mnuFormatIncrease 
         Caption         =   "����������(&I)"
      End
   End
   Begin VB.Menu mnuAct 
      Caption         =   "����(&A)"
      Begin VB.Menu mnuActReply 
         Caption         =   "��(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuActReplyAll 
         Caption         =   "ȫ����(&L)"
      End
      Begin VB.Menu mnuActForward 
         Caption         =   "ת��(&W)"
      End
      Begin VB.Menu mnuActSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActHigh 
         Caption         =   "��Ҫ�Ը�(&H)"
      End
      Begin VB.Menu mnuActLow 
         Caption         =   "��Ҫ�Ե�(&O)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMessageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���浱ǰ����༭����ϢID
Public mstrID As String     'strID ��ϢID�����Ϊ�գ���ʾ������Ϣ
Public mstrByID As String   'strByID Ҫ����ת�����𸴵�ԭʼID
Public mlng���� As Long     '��ǰ�û���Ӧ����Ϣ����
Public mlngMode As Long     '�򿪷�ʽ��1-�𸴣�2-ȫ���𸴣�3-ת����0-�½��ʼ�

Dim mstr�ỰID As String
Dim mblnSend As Boolean     '׼������
Dim mblnDelete As Boolean   '�Ƿ��Ѿ�ɾ��

Private mrsUser As ADODB.Recordset '����ռ�������,�û���,��Ա���ʵļ�¼��

Dim mlngFore As Long          '��ǰ��ǰ��ɫ

Dim mblnChange As Boolean
'�������ڲ��ҵ�
Public mblnCase As Boolean
Public mblnBegin As Boolean
Public mstrFind As String

Private Const EM_GETLINECOUNT = &HBA

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - 60
    
    '������һ�и��ؼ���λ��
    lbl(0).Top = sngTop + 120
    lbl(0).Left = 60
    lbl(1).Top = lbl(0).Top
    lbl(1).Left = ScaleWidth - lbl(1).Width - 60
    '�����ڶ��и��ؼ���λ��
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
    '���������и��ؼ���λ��
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
        
    '�����༭�ؼ���λ��
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
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    SaveWinState Me, App.ProductName
    
    mstr�ỰID = ""
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
    Dim str�ռ���  As String, strAddressee As String, strLastOne As String
    Dim rsUser As ADODB.Recordset
    Dim strData() As String
    Dim i As Long
    
    Set rsUser = mrsUser
    str�ռ��� = txtReceive.Text
    If frmSelectReceiver.Get�ռ���(str�ռ���, rsUser) = True Then
        '��"str�ռ���"���вü���ʹ������len(str�ռ���)<=txtReceive.maxlength������
        If Len(str�ռ���) > txtReceive.MaxLength Then
            strAddressee = Mid(str�ռ���, 1, txtReceive.MaxLength - 3)
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
            strAddressee = str�ռ���
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
'��
    frmMessageEdit.OpenWindow "", mstrID, mlng����, 1
    Unload Me
End Sub

Private Sub mnuActReplyAll_Click()
'ȫ����
    frmMessageEdit.OpenWindow "", mstrID, mlng����, 2
    Unload Me
End Sub

Private Sub mnuActForward_Click()
'ת����Ϣ
    frmMessageEdit.OpenWindow "", mstrID, mlng����, 3
    Unload Me
End Sub

Private Sub mnuActHigh_Click()
'������Ϣ����Ҫ�ԣ���
    mnuActHigh.Checked = Not mnuActHigh.Checked
    mnuActLow.Checked = False
    tlbNew.Buttons("High").Value = IIf(mnuActHigh.Checked, tbrPressed, tbrUnpressed)
    tlbNew.Buttons("Low").Value = IIf(mnuActLow.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuActLow_Click()
'������Ϣ����Ҫ�ԣ���
    mnuActHigh.Checked = False
    mnuActLow.Checked = Not mnuActLow.Checked
    tlbNew.Buttons("High").Value = IIf(mnuActHigh.Checked, tbrPressed, tbrUnpressed)
    tlbNew.Buttons("Low").Value = IIf(mnuActLow.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuEditCut_Click()
'���м���
    If rtfContent.SelLength = 0 Then Exit Sub
    
    Clipboard.SetText rtfContent.SelRTF, vbCFRTF
    rtfContent.SelText = ""
End Sub

Private Sub mnuEditCopy_Click()
'�����ı�
    If rtfContent.SelLength = 0 Then Exit Sub
    Clipboard.SetText rtfContent.SelRTF, vbCFRTF
End Sub

Private Sub mnuEditPaste_Click()
'ճ���ı�
    If Clipboard.GetFormat(vbCFText) = True Then
        If Clipboard.GetText(vbCFRTF) <> "" Then
            rtfContent.SelRTF = Clipboard.GetText(vbCFRTF)
        Else
            rtfContent.SelRTF = Clipboard.GetText(vbCFText)
        End If
    End If
End Sub

Private Sub mnuEditClear_Click()
'�����ѡ���ı�
    rtfContent.SelText = ""
End Sub

Private Sub mnuEditSel_Click()
'ȫ��ѡ��
    rtfContent.SelStart = 0
    rtfContent.SelLength = Len(rtfContent.Text)
End Sub

Private Sub mnuEditFind_Click()
'���ҵ�һ��
    Set frmMessageFind.frmMain = Me
    frmMessageFind.Show vbModal, Me
End Sub

Private Sub mnuEditFindNext_Click()
'��������
    Call FindText
End Sub

Private Sub mnuFileSave_Click()
'������Ϣ
    Dim lst As ListItem
    
    '����������
    If SaveMessage(False) = True Then
        '�����ж��Ƿ���ʵ���ʾλ��
        With frmMessageManager
            On Error Resume Next
            If mstrByID <> "" Then
                '������ת���ģ�����ԭ���ʼ���ͼ��
                Set lst = .lvwMain.ListItems("C" & mlng���� & mstrByID)
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
                'λ�ڲݸ�
                If .mlngIndex = 0 Then Exit Sub
            Else
                'λ����ɾ��
                If .mlngIndex = 3 Then Exit Sub
            End If
            Set lst = .lvwMain.ListItems.Add(, "C0" & mstrID, txtSubject.Text, "Script", "Script")
            If Err <> 0 Then
                Set lst = .lvwMain.ListItems("C0" & mstrID)
                Err.Clear
            End If
            On Error GoTo 0
            If mnuActHigh.Checked = True Or mnuActLow.Checked = True Then
                lst.SubItems(1) = IIf(mnuActHigh.Checked = True, "��", "��")
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
'������Ϣ
    Dim lst As ListItem
    
    If mnuFileSend.Enabled = False Then Exit Sub
    
    If SaveMessage(True) = True Then
        '����������
        
        '�����ж��Ƿ���ʵ���ʾλ��
        With frmMessageManager
            On Error Resume Next
            If mstrByID <> "" Then
                '����ԭ���ʼ���ͼ��
                Set lst = .lvwMain.ListItems("C" & mlng���� & mstrByID)
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
                'ɾ���ݸ���Ϣ
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
            '�����ѷ�����Ϣ
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
          '����ͼ��
          Call frmMessageRead.UpdateNotify
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
'���Ϊ�ļ�
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Filter = "RTF�ļ�(*.RTF)|*.rtf"
    '����ʱ����ʾ���Ҳ�����ֻ����
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
'��������ĸ������
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
'ѡ��������ɫ

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
'ѡ�񱳾���ɫ
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
'���ı�������Ŀ����
    mnuFormatSig.Checked = Not mnuFormatSig.Checked
    tlbEdit.Buttons("Sig").Value = IIf(mnuFormatSig.Checked, tbrPressed, tbrUnpressed)
    
    rtfContent.SelBullet = mnuFormatSig.Checked
End Sub

Private Sub mnuFormatLeft_Click()
'�ı�����
    mnuFormatLeft.Checked = True
    mnuFormatRight.Checked = False
    mnuFormatCenter.Checked = False
    tlbEdit.Buttons("Left").Value = tbrPressed
    
    rtfContent.SelAlignment = 0
End Sub

Private Sub mnuFormatCenter_Click()
'�ı�����
    mnuFormatLeft.Checked = False
    mnuFormatRight.Checked = False
    mnuFormatCenter.Checked = True
    tlbEdit.Buttons("Center").Value = tbrPressed
    
    rtfContent.SelAlignment = 2
End Sub

Private Sub mnuFormatRight_Click()
'�ı�����
    mnuFormatLeft.Checked = False
    mnuFormatRight.Checked = True
    mnuFormatCenter.Checked = False
    tlbEdit.Buttons("Right").Value = tbrPressed
    
    rtfContent.SelAlignment = 1
End Sub

Private Sub mnuFormatIncrease_Click()
' ����������
    Dim i As Long
    
    i = IIf(rtfContent.SelIndent < rtfContent.Width - 1000, rtfContent.SelIndent, rtfContent.Width - 1000)
    rtfContent.SelIndent = i + 360
End Sub

Private Sub mnuFormatDecrease_Click()
'����������
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
            
            'ѡ����ɫ
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
        'ѡ����ɫ
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
    '�������ݶ�̬�޸��ı���ĸ߶�
    Dim lngRet As Long
    Dim lngLastLine As Long                   '��������
    Dim lngLineHeight  As Long                        'ÿ�еĸ߶�
    
    Set Me.Font = txtReceive.Font
    lngLineHeight = Me.TextHeight("TT")
    lngRet = SendMessage(txtReceive.hWnd, EM_GETLINECOUNT, 0, 0&)
    If lngRet <> lngLastLine Then
        If txtReceive.Height + txtReceive.Top + lngLineHeight > Me.ScaleHeight And lngRet > 1 Then
            If lngLastLine <= lngRet - 1 Then
                Exit Sub '����Ѿ������߶ȣ�����
            End If
            lngLastLine = lngRet - 1    '�������߶�
        Else
            lngLastLine = lngRet
        End If
        txtReceive.Height = (lngLastLine + 1) * lngLineHeight     '�޸ĸ߶�
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
'���ݲ�ͬ��ѡ�����ݣ����¹������ϵ�״̬
    With rtfContent
        '����
        cmbFont.Text = IIf(IsNull(.SelFontName), "", .SelFontName)
        cmbSize.Text = IIf(IsNull(.SelFontSize), "", .SelFontSize)
        tlbEdit.Buttons("Bold").Value = IIf(IIf(IsNull(.SelBold), False, .SelBold), tbrPressed, tbrUnpressed)
        tlbEdit.Buttons("Italic").Value = IIf(IIf(IsNull(.SelItalic), False, .SelItalic), tbrPressed, tbrUnpressed)
        tlbEdit.Buttons("UnderLine").Value = IIf(IIf(IsNull(.SelUnderline), False, .SelUnderline), tbrPressed, tbrUnpressed)
        
        '����
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
        
        '��Ŀ����
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

Public Sub OpenWindow(ByVal strID As String, ByVal strByID As String, Optional ByVal lng���� As Long, Optional ByVal lngMode As Long)
'���ܣ����ݲ�������ʾ��Ϣ�༭����
'strID ��ϢID�����Ϊ�գ���ʾ������Ϣ
'strByID Ҫ����ת�����𸴵�ԭʼID
'lngMode  ����ʽ��1-�𸴣�2-ȫ���𸴣�3-ת��
    Dim frmMessage As frmMessageEdit
    Dim frmTemp As Form
    
    '���Ҹ���Ϣ�Ƿ��Ѿ����˱༭����
    If strID <> "" Then
        For Each frmTemp In Forms
            If frmTemp.Name = "frmMessageEdit" Then
                If frmTemp.mstrID = strID And frmTemp.mlng���� = lng���� Then
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
        frmMessage.mlng���� = lng����
        frmMessage.mlngMode = lngMode
    End If
    frmMessage.Show , gfrmMain
End Sub

Private Sub SetState()
    '�������ÿɼ���
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

    '�������ÿ�����
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
    
    'װ��������ֺ�
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
    
    '���ò˵�����
    mnuEditCut.Caption = "����(&T)" & vbTab & "Ctrl+X"
    mnuEditCopy.Caption = "����(&C)" & vbTab & "Ctrl+C"
    mnuEditPaste.Caption = "ճ��(&P)" & vbTab & "Ctrl+V"
    mnuEditClear.Caption = "���(&L)" & vbTab & "Del"
    mnuEditSel.Caption = "ȫѡ(&A)" & vbTab & "Ctrl+A"
    mnuFileSend.Caption = "����(&E)" & vbTab & "Ctrl+Enter"
    
    If InStr(frmMessageManager.mstrPrivs, "������Ϣ") = 0 Then
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
'װ������
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset, rsState As ADODB.Recordset
    Dim str������ As String
    Dim str�ռ��� As String
    Dim str����  As String
    
    On Error GoTo ErrH
    lngID = Val(IIf(mstrID = "", mstrByID, mstrID)) '
    If lngID = 0 Then
        mblnDelete = False
        mblnSend = True
        Exit Sub '��ȫ�µģ��ò���װ��
    End If
    
    '�õ��ʼ�����
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select A.* ,B.ɾ��,B.״̬ from zlmessages A,zlmsgState B " & _
        " where A.ID=B.��ϢID and B.��ϢID=[1] and b.����= [2] and B.�û�=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, mlng����, gstrDbUser)
    
    mstr�ỰID = rsTemp("�ỰID")
    mblnDelete = IIf(rsTemp("ɾ��") = 1, True, False)
    mblnSend = (mstrID = "") Or (mlng���� <> 2 And mlng���� <> 1) '���ʼ���δ�����ʼ�
    
    lbl(1).Caption = "����ʱ�䣺" & IIf(IsNull(rsTemp("ʱ��")), "", Format(rsTemp("ʱ��"), "yyyy-MM-dd HH:mm:ss"))
    str������ = IIf(IsNull(rsTemp("������")), "", rsTemp("������"))
    str�ռ��� = IIf(IsNull(rsTemp("�ռ���")), "", rsTemp("�ռ���"))
    txtSubject.Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    rtfContent.BackColor = IIf(IsNull(rsTemp("����ɫ")), RGB(255, 255, 255), rsTemp("����ɫ"))
    
    '�õ��ʵݵ�ַ
    gstrSQL = "select ����,�û�,��� from zlmsgstate where ��ϢID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)

    '�����յ��ռ��˼�¼��
    Set mrsUser = Rec.CopyNew(Nothing, True, , Array("�û���", adVarWChar, 30, Empty, "����", adVarWChar, 30, Empty, "�ռ���", adVarWChar, 30, Empty))

    '���½�������
    Select Case mlngMode
        Case 1 '��
            rsTemp.Filter = "����=0 or ����=1"
            mrsUser.AddNew
            mrsUser.Fields("����") = rsTemp("���")
            mrsUser.Fields("�û���") = rsTemp("�û�")
            txtReceive.Text = mrsUser.Fields("����")
            If Left(txtSubject.Text, 3) <> "�𸴣�" Then
                txtSubject.Text = "�𸴣�" & txtSubject.Text
            End If
        Case 2 'ȫ����
            str���� = ""
            If str�ռ��� = "������Ա" Or str�ռ��� = "��������Ա" Or str�ռ��� = "��������Ա" Then
                txtReceive.Text = str�ռ���
            ElseIf InStr(str�ռ���, "]") > 0 And InStr(str�ռ���, "[") > 0 Then
                txtReceive.Text = str�ռ���
            Else
                If str������ <> str�ռ��� And InStr(str�ռ���, str������ & ",") = 0 And InStr(str�ռ���, "," & str������) = 0 Then
                    txtReceive.Text = str������ & "," & str�ռ���
                Else
                    txtReceive.Text = str�ռ���
                End If
            End If
            If Left(txtSubject.Text, 3) <> "�𸴣�" Then
                txtSubject.Text = "�𸴣�" & txtSubject.Text
            End If
            rsTemp.Filter = "����=3 or ����=2"
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("����") = rsTemp("���")
                mrsUser.Fields("�û���") = rsTemp("�û�")
                
                str���� = str���� & rsTemp("���") & ","
                rsTemp.MoveNext

            Loop
            If str������ <> str���� And InStr(str����, str������ & ",") = 0 And InStr(str����, "," & str������) = 0 Then
                mrsUser.AddNew
                mrsUser.Fields("����") = str������
                mrsUser.Fields("�û���") = gstrDbUser
            End If

            
        Case 3 'ת��
            txtReceive.Text = ""
            If Left(txtSubject.Text, 3) <> "ת����" Then
                txtSubject.Text = "ת����" & txtSubject.Text
            End If
        Case Else
            txtReceive.Text = str�ռ���
            rsTemp.Filter = "����=3 or ����=2"
            
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("����") = rsTemp!���
                mrsUser.Fields("�û���") = rsTemp!�û�
                rsTemp.MoveNext
            Loop
    End Select
    
    
    lbl(0).Caption = "�����ˣ�     " & str������
    Me.Caption = "��Ϣ  " & IIf(txtSubject.Text = "", "", "-  " & txtSubject.Text)
    If mlngMode <> 0 Then
        '��ԭ����������
        With rtfContent
            .SelStart = 0
            .SelText = vbCrLf & "----------ԭʼ��Ϣ------------" & vbCrLf
            .SelStart = 2
            .SelLength = Len("----------ԭʼ��Ϣ------------")
            .SelFontName = "����"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelColor = 0
            
            .SelLength = Len(.Text)
            .SelIndent = 720
            
            .SelStart = 0
            .SelLength = 2
            .SelColor = RGB(0, 0, 255)
            .SelFontName = "����"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelLength = 0
        End With
    End If
    gstrSQL = "select ״̬ from zlmsgstate where ��ϢID=" & lngID & " and ����=" & mlng���� & " and �û�='" & gstrDbUser & "'"
    Set rsState = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsState, gstrSQL, Me.Caption)
    '�޸ı��
    gstrSQL = "Zl_Zlmsgstate_Edit(1," & lngID & "," & mlng���� & ",'" & gstrDbUser & "','" & gstrUserName & "',Null,'1' || substr( '" & rsState!״̬ & "',2))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call frmMessageRead.UpdateNotify
    Dim lst As ListItem
    On Error Resume Next
    '�ĳ��Ѷ�
    Set lst = frmMessageManager.lvwMain.ListItems("C" & mlng���� & lngID)
    If Not lst Is Nothing Then
        If lst.Icon <> "Script" Then '���ǲݸ����Ϣ�ſ��ܸ���
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
    Dim lng�ỰID As Long
    Dim lngCount As Long, i As Long
    Dim strSQL As String
    Dim strData() As String '�洢�û���������
    Dim blnTrans As Boolean
    
    'û�޸ģ���ֻ�Ǳ���
    
    If blnSend = True And txtReceive.Text = "" And mrsUser Is Nothing Then
        MsgBox "��ѡ���ռ��ˡ�", vbExclamation, gstrSysName
        cmdReceive.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtReceive.Text, txtReceive.MaxLength, cmdReceive.hWnd, "�ռ���") = False Then
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtSubject.Text, txtSubject.MaxLength, txtSubject.hWnd, "����") = False Then
        Exit Function
    End If
    If LenB(StrConv(rtfContent.TextRTF, vbFromUnicode)) > 4000 Then
        MsgBox "���ĵ��ַ�̫�࣬���߸�ʽ̫�����ˡ�", vbExclamation, gstrSysName
        rtfContent.SetFocus
        Exit Function
    End If
    
    If mstrID = "" Then
        lngID = zlDatabase.GetNextId("zlmessages")
        lng�ỰID = IIf(mstr�ỰID = "", lngID, mstr�ỰID)
    Else
        lngID = mstrID
        lng�ỰID = mstr�ỰID
    End If
    
    ReDim strData(0)
    On Error GoTo errHandle
    'ƴ���ռ�����Ϣ
    If Not mrsUser Is Nothing Then
        If mrsUser.State = adStateOpen Then
            If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
            Do Until mrsUser.EOF
                '���ռ����û��������������ַ�����ʽΪ���û���1,����1#�û���2,����2#�û���3,����3
                strSQL = IIf(strSQL = "", "", strSQL & "#") & mrsUser.Fields("�û���") & "," & mrsUser.Fields("����")
                '��Ϊ�����ַ�����gstrSQL������󳤶���4000������������ȹ�������ֶ��ύ
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
    '1�����ӻ��޸�zlmessages�еļ�¼��2�����ӻ��޸ķ����˵ļ�¼��3��ɾ�������ռ��˵ļ�¼��4��Ϊԭ�����ϴ𸴻�ת����־
    gstrSQL = "Zl_Zlmessages_New(" & lngID & _
                                "," & lng�ỰID & _
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
                                "," & mlng���� & ")"
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�����ռ��˵ļ�¼
    For i = 0 To UBound(strData)
        Call zlDatabase.ExecuteProcedure(strData(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mstrID = lngID
    mstr�ỰID = lng�ỰID
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
    '�����������Ϊ�գ�ֱ���˳�
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
        MsgBox "���ҽ�����û�ҵ���" & mstrFind & "��", vbInformation, gstrSysName
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
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

