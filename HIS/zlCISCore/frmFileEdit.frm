VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmFileEdit 
   Caption         =   "�����ļ�"
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
   StartUpPosition =   2  '��Ļ����
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
            Key             =   "Ԫ��"
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
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":0BF8
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":0E14
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileEdit.frx":1030
            Key             =   "ɾ��"
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
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "���没���ļ�"
               Object.Tag             =   "����"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "��ӡԤ������"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ����"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ʷ"
               Key             =   "��ʷ"
               Object.ToolTipText     =   "�鿴�����޶���ʷ"
               Object.Tag             =   "��ʷ"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "�༭"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "�༭"
               Object.ToolTipText     =   "ѡ����ȫ��ʾ��ģ��"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "Sample"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԫ��"
               Key             =   "Ԫ��"
               Description     =   "�༭"
               Object.ToolTipText     =   "ѡ��Ԫ��ʾ��ģ��"
               Object.Tag             =   "Ԫ��"
               ImageKey        =   "History"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "�ڵ�ǰԪ��֮ǰ�����µ�Ԫ��"
               Object.Tag             =   "����"
               ImageKey        =   "Insert"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Description     =   "�༭"
               Object.ToolTipText     =   "����ǰԪ�شӲ�����ɾȥ"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "���"
               Description     =   "�༭"
               Object.ToolTipText     =   "�ڲ���ĩβ����µ����ݣ��粡�̼�¼�������¼�ȣ�"
               Object.Tag             =   "���"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Description     =   "�༭"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "��������Ĳ����ı������"
               Object.Tag             =   "����"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "���ı��в��������ַ�"
               Object.Tag             =   "����"
               ImageKey        =   "SpecChar"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ı�"
               Key             =   "�ı�"
               Description     =   "�༭"
               Object.ToolTipText     =   "��ʾ�������ı�"
               Object.Tag             =   "�ı�"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ת��"
               Key             =   "ת��"
               Description     =   "�༭"
               Object.ToolTipText     =   "����ǰ������������ת�����ı�"
               Object.Tag             =   "ת��"
               ImageKey        =   "toText"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�༭"
               Key             =   "�༭"
               Description     =   "�༭"
               Object.ToolTipText     =   "�༭�������ͼ"
               Object.Tag             =   "�༭"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "���Ҳ��˲���"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
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
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPrintSet 
         Caption         =   "��ӡ����(&U)"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "�����&Excel"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParamSet 
         Caption         =   "��������(&M)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_History 
         Caption         =   "������ʷ(&H)"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "�����ı�(&C)"
      End
      Begin VB.Menu mnuEdit_Char 
         Caption         =   "�����ַ�(&S)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEdit_Text 
         Caption         =   "��ʾ�ı�(&D)"
      End
      Begin VB.Menu mnuEdit_Exchange 
         Caption         =   "ת���ı�(&T)"
      End
      Begin VB.Menu mnuEdit_Map 
         Caption         =   "�༭ͼ��(&G)"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "����(&A)"
      Begin VB.Menu mnuOrder_Add 
         Caption         =   "ȫ��ʾ��(&A)"
         Begin VB.Menu FileList 
            Caption         =   "��ʾ���ļ�"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuOrder_Demo 
         Caption         =   "Ԫ��ʾ��(&E)"
      End
      Begin VB.Menu mnuOrder_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrder_Imp 
         Caption         =   "������ʷ����(&H)"
      End
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrder_Insert 
         Caption         =   "����Ԫ��(&I)"
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "ɾ��Ԫ��(&D)"
      End
      Begin VB.Menu mnuOrder_Rec 
         Caption         =   "��Ӽ�¼(&R)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "���Ҳ���(&F)"
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "������Ϣ(&I)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewDiag 
         Caption         =   "�����ο�(&V)"
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "����ɸ��(&D)"
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public strPrivs As String       '�û����б�����ľ���Ȩ��

Private FileID As String
Private PatientID As String '����ID
Private CheckID As String '����ID��Һŵ�ID
Private PatientType As Integer '0=���ﲡ�� 1=סԺ����
Private FileTypeID As String '����ģ���ļ�ID
Private bSample As Boolean '�Ƿ�ʾ��
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1
Private FileType As Integer '��������
Private AdviceID As Long '���ҽ��ID
Private blnAllowEdit As Boolean

Private iCurrElementIndex As Integer '��ǰԪ��˳���
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Function ShowMe(ByVal sFileID As String, sPatientID As String, sCheckID As String, _
    iPatientType As Integer, sFileTypeID As String, bSampleFile As Boolean, frmParent As Object, Optional ByVal bAllowEdit As Boolean = True, Optional iFileType As Integer = 0, _
    Optional ByVal btModal As Byte = 0, Optional ByVal lngAdviceID As Long = 0) As Long
    Dim rsTmp As New ADODB.Recordset, i As Integer
    
    On Error Resume Next
    FileID = sFileID: PatientID = sPatientID: CheckID = sCheckID
    PatientType = iPatientType: FileTypeID = sFileTypeID: bSample = bSampleFile: AdviceID = lngAdviceID
    Me.Tag = FileID  '��Ÿô��ڱ༭�Ĳ�����¼ID
    
    iCurrElementIndex = 1

    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "����", "����", 900
        .Add , "����", "����", 900
    End With
    With Me.lvwItem
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    '��ȡ��ѡԪ���б�
    GetElementList
    mnuOrder_Add_FileList
    '��ȡ������Ϣ
    FileType = 0
    If bSample Then
        Me.Caption = "ȫ��ʾ��"
        stbInfo.Visible = False
    Else
        Me.Caption = "�����ļ�"
        stbInfo.Visible = True
        If Len(FileID) > 0 Then
            zlDatabase.OpenRecordset rsTmp, "Select ��������,�������� From ���˲�����¼ Where ID=" & FileID, Me.Name
        Else
            zlDatabase.OpenRecordset rsTmp, "Select ����,���� From �����ļ�Ŀ¼ Where ID=" & FileTypeID, Me.Name
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0): FileType = rsTmp(1)
        
        zlDatabase.OpenRecordset rsTmp, "Select Nvl(�����,0),Nvl(סԺ��,0),����,Nvl(�Ա�,' '),Nvl(����,' '),nvl(b.����,' ') As ����,nvl(c.����,' ') As ����,��ǰ���� From ������Ϣ a,���ű� b,���ű� c Where ����ID=" & PatientID & " And a.��ǰ����ID=b.ID(+) And a.��ǰ����ID=c.ID(+)", "zlCISCore"
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "�޲�����Ϣ"
        Else
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "����ţ�" & rsTmp(0), "סԺ�ţ�" & rsTmp(1))
                .Item(1).Text = "������" & rsTmp(2) & "���Ա�" & rsTmp(3) & "�����䣺" & rsTmp(4)
                If PatientType = 0 Then
                    .Item(2).Visible = False: .Item(3).Visible = False
                Else
                    .Item(2).Text = "���ң�" & rsTmp(5)
                    .Item(3).Text = "������" & rsTmp(6) & "�����ţ�" & NVL(rsTmp(7))
                End If
            End With
            
            Me.Caption = rsTmp(2) + "-" + Me.Caption
        End If
    End If
    
    ProFile1.AllowEdit = bAllowEdit: blnAllowEdit = bAllowEdit
    '����˵���������
    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.mnuEdit.Visible = bAllowEdit: Me.mnuOrder.Visible = bAllowEdit
    Me.mnuOrder_1.Visible = bAllowEdit
    For i = 1 To Me.tbrMain.Buttons.Count
        If Me.tbrMain.Buttons(i).Description = "�༭" Then Me.tbrMain.Buttons(i).Visible = bAllowEdit
    Next
    Select Case iFileType
        Case 4 '�������
            Me.mnuOrder_Insert.Visible = False
            Me.mnuOrder_Delete.Visible = False
            Me.mnuOrder_Rec.Visible = False
            Me.mnuOrder_2.Visible = False
            Me.tbrMain.Buttons("����").Visible = False
            Me.tbrMain.Buttons("ɾ��").Visible = False
            Me.tbrMain.Buttons("���").Visible = False
    End Select
    
    If bSample Then
        Me.mnuFile_History.Visible = False
        Me.mnuFile_4.Visible = False
        Me.tbrMain.Buttons("��ʷ").Visible = False
    Else
        If Len(FileID) = 0 Then
            Me.mnuFile_History.Enabled = False
        Else
            Me.mnuFile_History.Enabled = True
        End If
        Me.tbrMain.Buttons("��ʷ").Enabled = Me.mnuFile_History.Enabled
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
    If MsgBox("���ز���ʾ���󣬵�ǰ�������ݽ������ǣ��Ƿ������", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "���ڼ��ز�����"
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
    BeginShowProgress "���ڼ��ز�����"
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
    '---------Ȩ�޿���-------------
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
        If MsgBox("�Ƿ񱣴�༭�Ĳ���", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            mnuFileSave_Click
        End If
    End If
'    zlCommFun.OpenIme False
    
    Call SaveWinState(Me, App.ProductName)
    
    On Error Resume Next
    ParentForm.EditFile_UnLoad Me.hwnd  '�����ϼ����ڱ༭�ѹر�
    ProFile1.Release
End Sub

Private Sub lvwDemo_DblClick()
    Dim blnReadOnly As Boolean, i As Integer
    If Me.lvwDemo.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Me.lvwDemo.Tag
        Case "��ʷ"
            If lvwDemo.SelectedItem.Text = "����" Then
                FileID = Mid(lvwDemo.SelectedItem.Key, 2)
                blnReadOnly = Not blnAllowEdit
            Else
                FileID = Mid(lvwDemo.SelectedItem.Key, 2) * -1 '��ǰ�İ汾�ļ�¼ID�ø�����ʾ
                blnReadOnly = True
            End If
            ProFile1.AllowEdit = Not blnReadOnly
            Me.MousePointer = vbHourglass
            BeginShowProgress "���ڼ��ز�����"
            ProFile1.ShowFile FileID, PatientID, CheckID, PatientType, FileTypeID, bSample, , Me.prbRefresh, AdviceID
            ProFile1.SetActiveElement 1
            Me.prbRefresh.Visible = False
            Me.MousePointer = vbDefault
            Me.stbThis.Panels(2).Text = ""
            
            Me.lvwDemo.Visible = False
        
            '����˵���������
            Me.mnuFileSave.Visible = Not blnReadOnly: Me.mnuFileSplit(1).Visible = Not blnReadOnly
            Me.mnuEdit.Visible = Not blnReadOnly: Me.mnuOrder.Visible = Not blnReadOnly
            Me.mnuOrder_1.Visible = Not blnReadOnly
            For i = 1 To Me.tbrMain.Buttons.Count
                If Me.tbrMain.Buttons(i).Description = "�༭" Then Me.tbrMain.Buttons(i).Visible = Not blnReadOnly
            Next
        Case "��¼"
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
        BeginShowProgress "����ˢ�²�����"
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
    strSQL = "Select b.ID,b.�����ı�,a.��������,a.��д���� From ���˲�����¼ a,���˲������� b," + _
        "(Select b.Ԫ�ر���,Max(b.id) As ID From ���˲�����¼ a,���˲������� b,����Ԫ��Ŀ¼ c Where a.ID=b.������¼ID And b.Ԫ�ر���=c.���� And a.����id=" & PatientID & " And " + _
        IIf(PatientType = 1, "��ҳid=" & CheckID, "�Һŵ�='" & CheckID & "'") & " And (b.Ԫ������=0 Or c.���� Like 'ZL9CISCORE.%DIAG%') Group By b.Ԫ�ر���) c " + _
        "Where a.ID=b.������¼ID And b.ID=c.ID"
    strSQL = strSQL + " Union Select b.ID,b.�����ı�||'('||c.������Ŀ||')',a.��������,a.��д���� From ���˲�����¼ a,���˲������� b," + _
        "(Select b.Ԫ�ر���,nvl(d.����,' ') As ������Ŀ,Max(b.id) As ID From ���˲�����¼ a,���˲������� b,����Ԫ��Ŀ¼ c,���˲��������� d Where a.ID=b.������¼ID And b.Ԫ�ر���=c.���� And d.����id=b.id(+) And a.����id=" & PatientID & " And " + _
        IIf(PatientType = 1, "��ҳid=" & CheckID, "�Һŵ�='" & CheckID & "'") & " And c.���� Like 'ZL9CISCORE.%SPECRESULT%' And d.�ؼ���=-2 Group By b.Ԫ�ر���,nvl(d.����,' ')) c " + _
        "Where a.ID=b.������¼ID And b.ID=c.ID Order By ��д���� Desc"
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "�����ı�", True, , "��ѡ��ò�������Ĳ����ı�", , , True, _
        Me.Left + Me.tbrMain.Left + IIf(Me.cbrMain.Visible, Me.tbrMain.Buttons("����").Left, 0), Me.Top + Me.tbrMain.Top + 300 + IIf(Me.cbrMain.Visible, tbrMain.Buttons("����").Top + Me.tbrMain.Buttons("����").Height, 0), 0, , , True)
        
    If Not rsTmp Is Nothing Then
        lngContentID = rsTmp("ID"): rsTmp.Close
        Call zlDatabase.OpenRecordset(rsTmp, "Select a.ID,nvl(b.����,' ') From ���˲������� a,����Ԫ��Ŀ¼ b Where a.Ԫ�ر���=b.���� And a.ID=" & lngContentID, Me.Caption)
        
        If Not rsTmp.EOF Then ProFile1.CopyElement iCurrElementIndex, rsTmp("ID"), rsTmp(1)
    End If
End Sub

Private Sub mnuEdit_Exchange_Click()
    If MsgBox("���������ݽ��������ı������ݣ��Ƿ����", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    ProFile1.ChangeToText iCurrElementIndex
    
    If Not Me.mnuEdit_Text.Checked Then
        mnuEdit_Text_Click
    Else
        If Not ProFile1.ShowText(iCurrElementIndex, True) Then Me.mnuEdit_Text.Checked = False: Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuEdit_Map_Click()
    ProFile1.EditElement iCurrElementIndex
End Sub

Private Sub mnuEdit_Text_Click()
    If ProFile1.ShowText(iCurrElementIndex, Not Me.mnuEdit_Text.Checked) Then Me.mnuEdit_Text.Checked = Not Me.mnuEdit_Text.Checked
    Me.tbrMain.Buttons("�ı�").Value = IIf(Me.mnuEdit_Text.Checked, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile_History_Click()
    tbrMain_ButtonClick tbrMain.Buttons("��ʷ")
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
        FileID = sTmpFileID: Me.Tag = FileID '��Ÿô��ڱ༭�Ĳ�����¼ID
        SaveFile = True
    
        Me.mnuFile_History.Enabled = True
        Me.tbrMain.Buttons("��ʷ").Enabled = True
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
    Me.prbRefresh.Value = 0: BeginShowProgress "" '"����ˢ�²�����"
    ProFile1.DeleteElement iCurrElementIndex, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Demo_Click()
    tbrMain_ButtonClick tbrMain.Buttons("Ԫ��")
End Sub

Private Sub mnuOrder_Imp_Click()
    Dim lngImpId As Long    'Ҫ����Ĳ�����¼ID
    
    '��ȡ�����ļ�
    lngImpId = GetFileId(CLng(FileTypeID))
    
    If lngImpId = 0 Then Exit Sub
    If MsgBox("������ʷ�����󣬵�ǰ�������ݽ������ǣ��Ƿ������", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Err = 0: On Error Resume Next
    Me.MousePointer = vbHourglass
    BeginShowProgress "���ڼ��ز�����"
    ProFile1.LoadSample lngImpId, Me.prbRefresh, False
    ProFile1.SetActiveElement 1
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Insert_Click()
    tbrMain_ButtonClick tbrMain.Buttons("����")
End Sub

Private Sub mnuOrder_Rec_Click()
    tbrMain_ButtonClick tbrMain.Buttons("���")
End Sub

Private Sub mnuPreview_Click()
    Dim frmPreview As frmCasePrint
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If Len(FileID) = 0 Then
        If MsgBox("�ò����������ģ���ӡ֮ǰϵͳ������÷ݲ������Ƿ����", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1.Modified And ProFile1.AllowEdit Then _
            If MsgBox("��ӡ֮ǰ�Ƿ񱣴�÷ݲ���", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
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
        If MsgBox("�ò����������ģ���ӡ֮ǰϵͳ������÷ݲ������Ƿ����", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1.Modified And ProFile1.AllowEdit Then _
            If MsgBox("��ӡ֮ǰ�Ƿ񱣴�÷ݲ���", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
    End If
'            If MsgBox("׼����ӡ��������ӡ��׼��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(Me) Then
        MsgBox "��ӡ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.stbThis.Panels(2).Text = "�������ӡ�� " & Printer.DeviceName & " ���..."
    If bSample Then
        PrintOutCase Me, Printer, 0, True, 1, 0, FileID, False, 0, 1
    Else
        If 1 * FileID > 0 Then
            PrintOutCase Me, Printer, FileType, True, -1 * FileID, CLng(PatientID), CheckID, False, 0, 1
        Else
            PrintOutCase Me, Printer, FileType, True, 1, 0, CLng(FileID), False, 0, 1
        End If
    End If
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(Me)
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuPrintSet_Click()
    frmPrintSet.Show vbModal
End Sub

Private Sub mnuRefresh_Click()
    If MsgBox("�����������µ��뱣��Ĳ�������ǰ����" + Chr(13) + "���޸����δ���潫���������Ƿ������", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "���ڼ��ز�����"
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
        Case "Ԥ��"
            mnuPreview_Click
        Case "��ӡ"
            mnuPrint_Click
        Case "����"
            mnuFileSave_Click
        Case "����"
            With Me.lvwItem
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwDemo.Visible = False
                .SetFocus
            End With
        Case "ȫ��"
            Call PopupButtonMenu(Me.tbrMain, Button, Me.mnuOrder_Add)
        Case "��ʷ"
            With Me.lvwDemo
                GetFileHistory
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "Ԫ��"
            With Me.lvwDemo
                GetElementDemoList ProFile1.ElementID(iCurrElementIndex)
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "ɾ��"
            mnuOrder_Delete_Click
        Case "���"
            With Me.lvwDemo
                GetAddFile
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "����"
            mnuEdit_Copy_Click
        Case "����"
            mnuEdit_Char_Click
        Case "�ı�"
            mnuEdit_Text_Click
        Case "ת��"
            mnuEdit_Exchange_Click
        Case "�༭"
            mnuEdit_Map_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
            mnuExit_Click
    End Select
End Sub

Private Sub ShowEditMenu(ElementType As Integer)
    If Not ProFile1.AllowEdit Then Exit Sub
    Select Case ElementType
        Case 2 '������
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = True
            Me.tbrMain.Buttons("�ı�").Value = IIf(ProFile1.IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("ת��").Enabled = True
            Me.tbrMain.Buttons("����").Enabled = False 'True
            Me.tbrMain.Buttons("�༭").Enabled = False
        Case 3 '���ͼ
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = False
            Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
            Me.tbrMain.Buttons("ת��").Enabled = False
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�༭").Enabled = True
        Case 4 'ר��ֽ
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = True
            Me.tbrMain.Buttons("�ı�").Value = IIf(ProFile1.IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("ת��").Enabled = True
            Me.tbrMain.Buttons("����").Enabled = False 'True
            Me.tbrMain.Buttons("�༭").Enabled = False
        Case Else
            Me.tbrMain.Buttons("����").Enabled = IIf(ElementType = 0, True, False)
            Me.tbrMain.Buttons("�ı�").Enabled = False
            Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
            Me.tbrMain.Buttons("ת��").Enabled = False
            Me.tbrMain.Buttons("����").Enabled = IIf(ElementType = 0 Or ElementType = -5, True, False) 'True
            Me.tbrMain.Buttons("�༭").Enabled = False
    End Select
    
    Me.mnuEdit_Copy.Enabled = Me.tbrMain.Buttons("����").Enabled
    Me.mnuEdit_Char.Enabled = Me.tbrMain.Buttons("����").Enabled
    Me.mnuEdit_Map.Enabled = Me.tbrMain.Buttons("�༭").Enabled
    Me.mnuEdit_Text.Enabled = Me.tbrMain.Buttons("�ı�").Enabled
    Me.mnuEdit_Text.Checked = IIf(Me.tbrMain.Buttons("�ı�").Value = tbrPressed, True, False)
    Me.mnuEdit_Exchange.Enabled = Me.tbrMain.Buttons("ת��").Enabled
    
    Me.mnuViewDoctor.Visible = Not bSample
    
    If bSample Then
        Me.mnuEdit_Copy.Visible = False
        Me.tbrMain.Buttons("����").Visible = False
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
            gstrSql = "select I.ID,I.����,I.����,I.���� from ����Ԫ��Ŀ¼ I where substr(I.����,1,1)='1' And (����>=0 Or ����=-5) order by I.����"
        Case 1
            gstrSql = "select I.ID,I.����,I.����,I.���� from ����Ԫ��Ŀ¼ I where substr(I.����,2,1)='1' And (����>=0 Or ����=-5) order by I.����"
        Case 2
            gstrSql = "select I.ID,I.����,I.����,I.���� from ����Ԫ��Ŀ¼ I where substr(I.����,3,1)='1' And (����>=0 Or ����=-5) order by I.����"
        Case 3
            gstrSql = "select I.ID,I.����,I.����,I.���� from ����Ԫ��Ŀ¼ I where substr(I.����,4,1)='1' And (����>=0 Or ����=-5) order by I.����"
    End Select
    With rsTemp
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        If .BOF Or .EOF Then
            MsgBox "δ�����������Ƶ��ݵĲ���Ԫ�أ�", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Ԫ��": objItem.SmallIcon = "Ԫ��"
            objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
            strTemp = Switch(!���� = 0, "�ı���", !���� = 1, "���ӱ�", !���� = 2, "������", !���� = 3, "���ͼ", !���� = 4, "ר��ֽ", _
                            !���� = -1, "��дǩ��", !���� = -2, "��ǰ����", !���� = -3, "��ǰʱ��", !���� = -4, "�������", !���� = -5, "��ͨ�ı�")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = strTemp
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
    '����ļ��嵥
    iNum = FileList.Count
    FileList(0).Visible = True
    For i = 1 To iNum - 1
        Unload FileList(i)
    Next
    
    If Len(FileTypeID) = 0 Then
        If bSample Then
            zlDatabase.OpenRecordset rsFileList, "Select �ļ�ID From" + _
            " ����ʾ��Ŀ¼ Where ID=" & FileID, Me.Caption
            
            FileTypeID = rsFileList(0)
        Else
            zlDatabase.OpenRecordset rsFileList, "Select �ļ�ID From" + _
            " ���˲�����¼ Where ID=" & FileID, Me.Caption
            
            FileTypeID = rsFileList(0)
        End If
    End If
    
    zlDatabase.OpenRecordset rsFileList, "Select a.ID,a.���� From ����ʾ��Ŀ¼ a" + _
        " Where a.�ļ�ID=" & FileTypeID & " And a.����=1" + _
        IIf(bSample, " And a.ID<>" & FileID, "") + _
        IIf(bSample, "", " And (a.����ID=" & UserInfo.����ID & " Or" + _
        " a.����ID Is Null)"), Me.Caption
    If rsFileList.EOF Then Exit Sub
    
    i = 1
    Do While Not rsFileList.EOF
        Load FileList(FileList.Count)
        With FileList(FileList.Count - 1)
            .Caption = "&" & i & " " & rsFileList("����")
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
    zlDatabase.OpenRecordset rsTemp, "Select a.ID,a.����,a.˵�� From ����ʾ��Ŀ¼ a" + _
        " Where a.Ԫ��ID=" & ElementID & " And a.����=2" + _
        IIf(bSample, "", " And (a.����ID=" & UserInfo.����ID & " Or" + _
        " a.����ID Is Null)"), Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = ""
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "˵��", "˵��", 1800
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Ԫ��": objItem.SmallIcon = "Ԫ��"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("˵��").Index - 1) = IIf(IsNull(!˵��), "", !˵��)
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
'��ȡ�����޶���ʷ
Private Sub GetFileHistory()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    If Len(FileID) = 0 Then Exit Sub
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo errHand
    zlDatabase.OpenRecordset rsTemp, "Select '����' As �汾,������ As ��д��,�������� As ��д����,ID From ���˲�����¼ Where ID=" & Me.Tag & _
        " Union All Select to_Char(�汾���,'9999') As �汾,��д��,��д����,ID From ���˲����޶���¼ Where ������¼ID=" & Me.Tag & _
        " Order By �汾 Desc", Me.Caption
    
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = "��ʷ"
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "�汾", "�汾", 800
        .Add , "��д��", "��д��", 1000
        .Add , "ʱ��", "ʱ��", 1800
    End With
    With Me.lvwDemo
        .ColumnHeaders("�汾").Position = 1
        .SortKey = .ColumnHeaders("�汾").Index - 1
        .SortOrder = lvwDescending
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !�汾)
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("��д��").Index - 1) = IIf(IsNull(!��д��), "", !��д��)
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("ʱ��").Index - 1) = IIf(IsNull(!��д����), "", !��д����)
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
'��ȡ�ɸ��ӵĲ����嵥
Private Sub GetAddFile()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As MSComctlLib.ListItem
    Dim strTemp As String
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo errHand
    zlDatabase.OpenRecordset rsTemp, "Select a.ID,a.����,a.˵�� From �����ļ�Ŀ¼ a" + _
        " Where a.����=" & FileType & " And Nvl(a.����,0)=1", Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    Me.lvwDemo.Tag = "��¼"
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "˵��", "˵��", 1800
    End With
    
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Ԫ��": objItem.SmallIcon = "Ԫ��"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("˵��").Index - 1) = IIf(IsNull(!˵��), "", !˵��)
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
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

