VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmBillEdit 
   Caption         =   "���Ƶ���"
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
   StartUpPosition =   3  '����ȱʡ
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
            Key             =   "Ԫ��"
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
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1CC6
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":1EE2
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":20FE
            Key             =   "ɾ��"
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
               Caption         =   "����"
               Key             =   "����"
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
               Object.Visible         =   0   'False
               Description     =   "�༭1"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "�༭1"
               Object.ToolTipText     =   "ѡ����ȫ��ʾ��ģ��"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "Sample"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Ԫ��"
               Key             =   "Ԫ��"
               Description     =   "�༭1"
               Object.ToolTipText     =   "ѡ��Ԫ��ʾ��ģ��"
               Object.Tag             =   "Ԫ��"
               ImageKey        =   "History"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭1"
               Object.ToolTipText     =   "�ڵ�ǰԪ��֮ǰ�����µ�Ԫ��"
               Object.Tag             =   "����"
               ImageKey        =   "Insert"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Description     =   "�༭1"
               Object.ToolTipText     =   "����ǰԪ�شӲ�����ɾȥ"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Description     =   "�༭"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "��������Ĳ����ı������"
               Object.Tag             =   "����"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "�༭"
               Object.ToolTipText     =   "���ı��в��������ַ�"
               Object.Tag             =   "����"
               ImageKey        =   "SpecChar"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ı�"
               Key             =   "�ı�"
               Description     =   "�༭"
               Object.ToolTipText     =   "��ʾ�������ı�"
               Object.Tag             =   "�ı�"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ת��"
               Key             =   "ת��"
               Description     =   "�༭"
               Object.ToolTipText     =   "����ǰ������������ת�����ı�"
               Object.Tag             =   "ת��"
               ImageKey        =   "toText"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�༭"
               Key             =   "�༭"
               Description     =   "�༭"
               Object.ToolTipText     =   "�༭�������ͼ"
               Object.Tag             =   "�༭"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Description     =   "���"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "���"
               Description     =   "���"
               Object.ToolTipText     =   "��˵�ǰ����"
               Object.Tag             =   "���"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "���ص�ǰ����"
               Object.Tag             =   "����"
               ImageKey        =   "Rollback"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_41"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "���Ҳ��˲���"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ʾ"
               Key             =   "��ʾ"
               Object.ToolTipText     =   "��ʾ����ģ��"
               Object.Tag             =   "��ʾ"
               ImageKey        =   "History"
               Style           =   1
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ģ��"
               Key             =   "ģ��"
               Object.ToolTipText     =   "����ǰ�ı����ݱ���Ϊ����ģ��"
               Object.Tag             =   "ģ��"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillEdit.frx":B6E4
            Key             =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         Begin VB.TextBox txt�ɼ� 
            Height          =   300
            Left            =   5040
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmd�ɼ� 
            Height          =   285
            Left            =   6600
            Picture         =   "frmBillEdit.frx":D5E0
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "ѡ�����걾"
            Top             =   350
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   6440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chk��ʼʱ�� 
            BackColor       =   &H80000005&
            Caption         =   "Ҫ��ʱ��"
            Height          =   225
            Left            =   315
            TabIndex        =   4
            ToolTipText     =   "�Ƿ���ʱ��"
            Top             =   420
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txt���� 
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
         Begin VB.TextBox txtƵ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1350
            TabIndex        =   10
            Top             =   1080
            Width           =   2500
         End
         Begin VB.TextBox txt���� 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4725
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1080
            Width           =   1380
         End
         Begin VB.CheckBox chk���� 
            BackColor       =   &H80000005&
            Caption         =   "����(&J)"
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
            ToolTipText     =   "ѡ�����걾"
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   285
            Left            =   5280
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cboִ�п��� 
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
         Begin VB.TextBox txtҽ������ 
            Height          =   300
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   3945
         End
         Begin VB.ComboBox cboҽ�� 
            Height          =   300
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1425
            Width           =   1590
         End
         Begin VB.TextBox txtҽ������ 
            Height          =   300
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   9
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdƵ�� 
            Enabled         =   0   'False
            Height          =   240
            Left            =   3575
            Picture         =   "frmBillEdit.frx":D7D0
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(F4)"
            Top             =   1110
            Width           =   270
         End
         Begin MSComCtl2.DTPicker txt��ʼʱ�� 
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
         Begin VB.Label lbl�ɼ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɼ���ʽ"
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����걾"
            Height          =   180
            Left            =   5640
            TabIndex        =   39
            Top             =   45
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ��"
            Height          =   180
            Left            =   6660
            TabIndex        =   37
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   8460
            TabIndex        =   36
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lblƵ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ƶ��"
            Height          =   180
            Left            =   960
            TabIndex        =   35
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lbl������λ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   6150
            TabIndex        =   34
            Top             =   1140
            Width           =   15
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   4335
            TabIndex        =   33
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lblִ�п��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�п���"
            Height          =   180
            Left            =   600
            TabIndex        =   32
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ŀ"
            Height          =   180
            Left            =   600
            TabIndex        =   31
            Top             =   45
            Width           =   720
         End
         Begin VB.Label lbl��ʼʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ҫ��ʱ��"
            Height          =   180
            Left            =   600
            TabIndex        =   30
            Top             =   435
            Width           =   720
         End
         Begin VB.Label lbl����ҽ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽ��"
            Height          =   180
            Left            =   5175
            TabIndex        =   29
            Top             =   1485
            Width           =   720
         End
         Begin VB.Label lblҽ������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ������"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEdit_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Auditing 
         Caption         =   "��˱���(&A)"
      End
      Begin VB.Menu mnuEdit_Rollback 
         Caption         =   "���ر���(&B)"
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
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "�������(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Template 
         Caption         =   "����ģ��(&M)"
      End
   End
   Begin VB.Menu mnuOrder_1 
      Caption         =   "����(&A)"
      Visible         =   0   'False
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
      Begin VB.Menu mnuOrder_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Insert 
         Caption         =   "����Ԫ��(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrder_Delete 
         Caption         =   "ɾ��Ԫ��(&D)"
         Visible         =   0   'False
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
      Begin VB.Menu mnuTemplate 
         Caption         =   "����ģ��(&M)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPatientInformation 
         Caption         =   "������Ŀ(&I)"
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
Attribute VB_Name = "frmBillEdit"
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
Private bln��ʿվ As Boolean
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1
Private DeptID As Long '��������
Private mblnShow���� As Boolean
Private PatientDate As Date '���˾������Ժʱ��
Private AdviceID As Long, SendNO As Long 'ҽ��ID�����ͺ�
Private sCheckNo As String '���͵��ݺ�
Private iRecordType As Integer '��¼����
Private alngFileID(1) As Long '����ͱ���ID
Private intType As Integer '�������:-1=������0=�����ϡ�1=������2=��ҩ��4=����
Private iTabIndex As Integer
Private mlngǰ��ID As Long, blnҽ��ִ�� As Boolean
Private mblnMoved As Boolean
Private mstrPrivs As String

'ҽ���༭
Private strAdviceText As String 'ҽ������
Private str��� As String, lngClinicID As Long, strClinicName As String, str�걾��λ As String
Private strSequence As String, lngƵ�ʴ��� As Long, lngƵ�ʼ�� As Long, str�����λ As String 'Ƶ��
Private int�Ƽ����� As Integer, intִ������ As Integer, lng���˿���ID As Long
Private mstr�Ա� As String
Private mstrLike As String
Private gint�����Ǽ���Ч���� As Integer
Private rsRelativeAdvice As ADODB.Recordset '���ҽ��
Private strExtData As String '������Ŀ

Private ifInitItem As Boolean '�Ƿ��ڽ�������ʱֱ����ʾ������Ŀ

Private iCurrElementIndex As Integer '��ǰԪ��˳���
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Sub ShowMe(frmParent As Object, ByVal lngҽ��ID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strҽ������ As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional ByVal blnMoved As Boolean = False)
'strPrivs��Ȩ�޴���ÿһλ����һ��Ȩ�ޣ�0���޸�Ȩ�ޡ�1���и�Ȩ�ޡ�
'   ��1λ����˱���
'   ��2λ�����ر���
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String, tmpDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    Dim rsDept As New ADODB.Recordset, strDept As String, strDeptName As String
    Dim strSQL As String
    
    On Error Resume Next
    '��ʼ��
    If blnMoved Then ReadOnly = True
    mblnMoved = blnMoved
    
    strSQL = "Select a.����ID,a.��ҳID,a.�Һŵ�,Decode(a.��ҳID,Null,0,1),b.ID,b.����,a.ҽ������," + _
        "ҽ������,��ʼִ��ʱ��,������־,ִ��Ƶ��,�ܸ�����,��������,����ҽ��,nvl(b.���㵥λ,' ') As ���㵥λ,b.���,nvl(a.�걾��λ,' ') As �걾��λ,A.ִ�п���ID " + _
        "From ����ҽ����¼ a,������ĿĿ¼ b Where (a.ID=[1] Or a.���ID=[1]) And a.������ĿID=b.ID Order By nvl(a.���ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, lngҽ��ID)
    If rsTmp.EOF Then Unload Me: Exit Sub
    lngClinicID = rsTmp(4): strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    '���츽����Ŀ��
    rsTmp.MoveNext
    If Not rsTmp.EOF Then
        If rsTmp!��� = "C" Then lngClinicID = rsTmp(4) '������Ŀ
    End If
    Do While Not rsTmp.EOF
        strExtData = strExtData & "," & rsTmp(4)
        If rsTmp!��� = "C" Then tmpDiagName = tmpDiagName & "," & rsTmp(5)
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    If Len(tmpDiagName) > 0 Then '������Ŀ
        strDiagName = Mid(tmpDiagName, 2)
        
        '�òɼ���ʽ
        rsTmp.MoveFirst
        Me.cmd�ɼ�.Tag = rsTmp(4)
        Me.txt�ɼ� = rsTmp(5): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        
        rsTmp.MoveNext
    Else
        rsTmp.MoveFirst
    End If
    
    intType = -1
    Me.txtҽ������ = strDiagName
    If rsTmp!��� = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intType = 0
        Call AdviceSet�������(1, strExtData)
        txtҽ������.Text = Get�����������(1, strDiagName)
        Me.txt���� = Get��λ����
    ElseIf rsTmp!��� = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intType = 1
        Call AdviceSet�������(2, strExtData)
        txtҽ������.Text = Get�����������(2, strDiagName)
        Me.txt���� = Get��������
    ElseIf InStr(",7,8,", rsTmp!���) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intType = 2
    ElseIf rsTmp!��� = "C" Then
        '������Ŀѡ�����걾
        intType = 4
        Me.txt���� = rsTmp("�걾��λ"): str�걾��λ = rsTmp("�걾��λ")
        strExtData = strExtData & ";" & str�걾��λ
    End If
    
    alngFileID(0) = lng����ID: PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    PatientType = rsTmp(3): FileTypeID = lng����ID: bSample = False: AdviceID = lngҽ��ID
    
    '��ʾҽ������
    If IsNull(rsTmp("��ʼִ��ʱ��")) Then
        Me.chk��ʼʱ��.Visible = True: Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"): Me.txt��ʼʱ��.Enabled = False
    Else
        Me.txt��ʼʱ�� = rsTmp("��ʼִ��ʱ��"): Me.txt��ʼʱ��.Enabled = True
    End If
    Me.chk����.Value = rsTmp("������־")
    If Not IsNull(rsTmp("ҽ������")) Then Me.txtҽ������ = rsTmp("ҽ������")
    Me.txtƵ�� = rsTmp("ִ��Ƶ��"): Me.txtƵ��.Enabled = True: Me.cmdƵ��.Enabled = True
    Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    If Not IsNull(rsTmp("�ܸ�����")) Then Me.txt���� = rsTmp("�ܸ�����"): Me.txt����.Enabled = True
    If Not IsNull(rsTmp("��������")) Then Me.txt���� = rsTmp("��������"): Me.txt����.Enabled = True: Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    
    Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False
    strSQL = "Select ����,���� From ���ű� Where ID=[1]"
    Set rsDept = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("ִ�п���ID"), 0))
    If Not rsDept.EOF Then
        Me.cboִ�п���.AddItem rsDept("����") & "-" & rsDept("����")
        Me.cboִ�п���.Text = rsDept("����") & "-" & rsDept("����"): Me.cboִ�п���.Enabled = True
    End If
    Me.cboҽ��.Clear: Me.cboҽ��.AddItem rsTmp("����ҽ��")
    Me.cboҽ��.Text = rsTmp("����ҽ��"): Me.cboҽ��.Enabled = True
    Me.picAdvice.Enabled = False
    
    Me.stbThis.Panels(3).Visible = False: Me.stbThis.Panels(4).Visible = False
    
    If alngFileID(0) = 0 Then
        strSQL = "Select Count(*)" + _
            " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=1"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            MsgBox "δ����������Ŀ�����ܱ༭", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From ���˲������� Where ������¼ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From �����ļ���� Where �����ļ�ID=" + FileTypeID + " And ��дʱ��=1"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    MsgBox "δ����������Ŀ�����ܱ༭", vbInformation, gstrSysName
                    Unload Me
                    Exit Sub
                End If
            Else
                MsgBox "û���������ݣ����ܱ༭", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        End If
    End If
    '��ʼ������
    
    '�ж��ܷ�༭����
    If Not ReadOnly Then
        '�˴��϶�����ѯ�󱸱�
        strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID=[1] And Not ����ID Is Null"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, AdviceID)
        If Not rsTmp.EOF Then ReadOnly = True
    End If
    bAllowEdit = Not ReadOnly
    
    iCurrElementIndex = 0

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
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "˵��", "˵��", 1800
    End With

    '��ȡ��ѡԪ���б�
'    GetElementList
'    mnuOrder_Add_FileList
    '��ȡ������Ϣ
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If bSample Then
        Me.Caption = "ȫ��ʾ��"
        stbInfo.Visible = False
    Else
        Me.Caption = "�����ļ�(����)"
        stbInfo.Visible = True
        If alngFileID(0) > 0 Then
            strSQL = "Select �������� From ���˲�����¼ Where ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        Else
            strSQL = "Select ���� From �����ļ�Ŀ¼ Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(����)"
        
        strSQL = "Select Nvl(�����,0),Nvl(סԺ��,0),����,Nvl(�Ա�,' '),Nvl(����,' '),nvl(b.����,' ') As ����,nvl(c.����,' ') As ����,��ǰ����," + IIf(PatientType = 0, "����ʱ�� ", "��Ժʱ�� ") + _
            "From ������Ϣ a,���ű� b,���ű� c Where ����ID=[1] And a.��ǰ����ID=b.ID(+) And a.��ǰ����ID=c.ID(+)"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "�޲�����Ϣ"
        Else
            PatientDate = rsTmp(8)
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "����ţ�" & rsTmp(0), "סԺ�ţ�" & rsTmp(1))
                .Item(1).Text = "������" & rsTmp(2) & "���Ա�" & rsTmp(3) & "�����䣺" & rsTmp(4)
                
                mstr�Ա� = rsTmp(3)
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
'    With Me.stbAdvInfo.Panels
'        .Item(1).Text = "��Ŀ��" + strDiagName
'        .Item(2).Text = "ҽ�����ݣ�" + strҽ������
'    End With
'    Me.stbDrAdviceInfo.Panels(1).Text = "ҽ�����У�" + strDrAdvice
    
    ProFile1(0).AllowEdit = bAllowEdit
    '����˵���������
    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.tbrMain.Buttons("����").Visible = bAllowEdit
    Me.mnuEdit_Clear.Visible = bAllowEdit
    
    iTabIndex = -1
    TabFile.Tabs.Clear
    TabFile.Tabs.Add , "����", "����(&S)", "����"
    TabFile.Tabs("����").Selected = True
    '����Tab
    Me.TabFile.Visible = False: Me.ProFile1(1).Visible = False

    Set ParentForm = frmParent
    
    SetItemFormat
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
End Sub

Public Sub ShowMe_Report(frmParent As Object, ByVal strNO As String, ByVal int��¼���� As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strҽ������ As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional ByVal lngǰ��ID As Long = 0, Optional ByVal Ifҽ��ִ�� As Boolean = False, Optional ByVal blnShow���� As Boolean = True, Optional ByVal lngҽ��ID As Long = 0, Optional blnMoved As Boolean = False, Optional strPrivs As String = "00")
    
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String, tmpDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    Dim rsDept As New ADODB.Recordset, strDept As String, strDeptName As String
    Dim rsCapture As New ADODB.Recordset '�ɼ���ʽ��¼
    Dim strSQL As String
    
    On Error Resume Next
    '��ʼ��
    If blnMoved Then ReadOnly = True
    mblnMoved = blnMoved
    mstrPrivs = strPrivs
    
    If ReadOnly Then
        tvwElement.Visible = False
        tbrMain.Buttons("��ʾ").Visible = False
        tbrMain.Buttons("Split_5").Visible = False
        mnuTemplate.Visible = False
    End If
    
    mblnShow���� = blnShow����
    
    picAdvice.Visible = mnuPatientInformation.Checked
'    If blnShow���� = False Then tvwElement.Visible = blnShow����
    
    strSQL = "Select a.����ID,a.��ҳID,a.�Һŵ�,Decode(a.��ҳID,Null,0,1),b.ID,b.����,a.ҽ������,a.ID,a.����ID," + _
        "ҽ������,��ʼִ��ʱ��,������־,ִ��Ƶ��,�ܸ�����,��������,����ҽ��,b.���,nvl(a.�걾��λ,' ') As �걾��λ,c.���ͺ�,ִ�п���ID,a.���ID " + _
        "From ����ҽ����¼ a,������ĿĿ¼ b,����ҽ������ c Where" & _
        " c.NO=[1] And c.��¼����=[2]" & _
        IIf(lngҽ��ID = 0, "", " And (A.ID=[3] Or A.���ID=[3])") & " And a.������ĿID=b.ID And a.ID=c.ҽ��ID Order By nvl(a.���ID,0)"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, strNO, int��¼����, lngҽ��ID)
    If rsTmp.EOF Then Unload Me: Exit Sub
    lngClinicID = rsTmp(4): strDiagName = rsTmp(5): strDrAdvice = rsTmp(6)
    
    sCheckNo = strNO: iRecordType = int��¼����
        
    '���츽����Ŀ��
'    If Not rsTmp!��� = "C" Then rsTmp.MoveNext
    Do While Not rsTmp.EOF
        If rsTmp!��� = "C" Then
            tmpDiagName = tmpDiagName & "," & rsTmp(5)
            strExtData = strExtData & "," & rsTmp(4)
        End If
    
        rsTmp.MoveNext
    Loop
    If Len(strExtData) > 0 Then strExtData = Mid(strExtData, 2)
    rsTmp.MoveFirst
    If Len(tmpDiagName) > 0 Then '������Ŀ
        strDiagName = Mid(tmpDiagName, 2)
        If Not rsTmp!��� = "C" Then rsTmp.MoveNext
        
        '�òɼ���ʽ
        strSQL = "Select b.ID,b.���� From ����ҽ����¼ a ,������ĿĿ¼ b " & _
            "Where a.������ĿID=b.ID and a.id=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsCapture = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("���ID"), 0))
        If Not rsCapture.EOF Then
            Me.cmd�ɼ�.Tag = rsCapture(0)
            Me.txt�ɼ� = rsCapture(1): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        End If
    End If
     
    intType = -1
    Me.txtҽ������ = strDiagName
    If rsTmp!��� = "D" And zlCommFun.NVL(GetItemField(rsTmp(4), "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intType = 0
        Call AdviceSet�������(1, strExtData)
        txtҽ������.Text = Get�����������(1, strDiagName)
        Me.txt���� = Get��λ����
    ElseIf rsTmp!��� = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intType = 1
        Call AdviceSet�������(2, strExtData)
        txtҽ������.Text = Get�����������(2, strDiagName)
        Me.txt���� = Get��������
    ElseIf InStr(",7,8,", rsTmp!���) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intType = 2
    ElseIf rsTmp!��� = "C" Then
        '������Ŀѡ�����걾
        intType = 4
        Me.txt���� = rsTmp("�걾��λ"): str�걾��λ = rsTmp("�걾��λ")
        strExtData = strExtData & ";" & str�걾��λ
    End If
   
    alngFileID(0) = IIf(IsNull(rsTmp(8)), 0, rsTmp(8))
    alngFileID(1) = lng����ID: PatientID = rsTmp(0): CheckID = IIf(rsTmp(3) = 0, rsTmp(2), rsTmp(1))
    
    PatientType = rsTmp(3): FileTypeID = lng����ID: bSample = False: AdviceID = rsTmp(7): SendNO = rsTmp("���ͺ�")
    mlngǰ��ID = lngǰ��ID: blnҽ��ִ�� = Ifҽ��ִ��
    
    '��ʾҽ������
    If IsNull(rsTmp("��ʼִ��ʱ��")) Then
        Me.chk��ʼʱ��.Visible = True: Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"): Me.txt��ʼʱ��.Enabled = False
    Else
        Me.txt��ʼʱ�� = rsTmp("��ʼִ��ʱ��"): Me.txt��ʼʱ��.Enabled = True
    End If
    Me.chk����.Value = rsTmp("������־")
    If Not IsNull(rsTmp("ҽ������")) Then Me.txtҽ������ = rsTmp("ҽ������")
    Me.txtƵ�� = rsTmp("ִ��Ƶ��"): Me.txtƵ��.Enabled = True: Me.cmdƵ��.Enabled = True
    Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    If Not IsNull(rsTmp("�ܸ�����")) Then Me.txt���� = rsTmp("�ܸ�����"): Me.txt����.Enabled = True
    If Not IsNull(rsTmp("��������")) Then Me.txt���� = rsTmp("��������"): Me.txt����.Enabled = True: Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = Trim(rsTmp("���㵥λ"))
    
    Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False
    strSQL = "Select ����,���� From ���ű� Where ID=[1]"
    Set rsDept = OpenSQLRecord(strSQL, Me.Caption, NVL(rsTmp("ִ�п���ID"), 0))
    If Not rsDept.EOF Then
        Me.cboִ�п���.AddItem rsDept("����") & "-" & rsDept("����")
        Me.cboִ�п���.Text = rsDept("����") & "-" & rsDept("����"): Me.cboִ�п���.Enabled = True
    End If
    Me.cboҽ��.Clear: Me.cboҽ��.AddItem rsTmp("����ҽ��")
    Me.cboҽ��.Text = rsTmp("����ҽ��"): Me.cboҽ��.Enabled = True
    Me.picAdvice.Enabled = False
    
    Me.stbThis.Panels(3).Text = "�����ˣ�" + UserInfo.����: Me.stbThis.Panels(4).Text = "ʱ�䣺" + Format(zlDatabase.Currentdate, "yy-MM-dd HH:mm:ss")
    
    If alngFileID(0) = 0 Then
        strSQL = "Select Count(*)" + _
            " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=1"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            alngFileID(0) = -1 'û��������Ŀ
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From ���˲������� Where ������¼ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(0))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=1"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    alngFileID(0) = -1 'û��������Ŀ
                End If
            Else
                alngFileID(0) = -1 'û��������Ŀ
            End If
        End If
    End If
    
    If alngFileID(1) = 0 Then
        strSQL = "Select Count(*)" + _
            " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=2"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        If rsTmp(0) = 0 Then
            MsgBox "δ���屨����Ŀ�����ܱ༭", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From ���˲������� Where ������¼ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "���˲�������", "H���˲�������")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(1))
        If rsTmp(0) = 0 Then
            If Len(FileTypeID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=2"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
                If rsTmp(0) = 0 Then
                    MsgBox "δ���屨����Ŀ�����ܱ༭", vbInformation, gstrSysName
                    Unload Me
                    Exit Sub
                End If
            Else
                MsgBox "û�б������ݣ����ܱ༭", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        End If
    End If
    '��ʼ������
    
    '�ж��ܷ�༭����
    If Not ReadOnly Then
        strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID=[1] And Not ����ID Is Null"
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
        .Add , "����", "����", 1800
        .Add , "����", "����", 900
        .Add , "����", "����", 900
    End With
    With Me.lvwItem
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "˵��", "˵��", 1800
    End With

    '��ȡ��ѡԪ���б�
'    GetElementList
'    mnuOrder_Add_FileList
    '��ȡ������Ϣ
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If bSample Then
        Me.Caption = "ȫ��ʾ��"
        stbInfo.Visible = False
    Else
        Me.Caption = "�����ļ�(����)"
        stbInfo.Visible = True
        If alngFileID(1) > 0 Then
            strSQL = "Select �������� From ���˲�����¼ Where ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, alngFileID(1))
        Else
            strSQL = "Select ���� From �����ļ�Ŀ¼ Where ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
        End If
        If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(����)"
        
        strSQL = "Select Nvl(�����,0),Nvl(סԺ��,0),����,Nvl(�Ա�,' '),Nvl(����,' '),nvl(b.����,' ') As ����,nvl(c.����,' ') As ����,��ǰ����," + IIf(PatientType = 0, "����ʱ�� ", "��Ժʱ�� ") + _
            "From ������Ϣ a,���ű� b,���ű� c Where ����ID=[1] And a.��ǰ����ID=b.ID(+) And a.��ǰ����ID=c.ID(+)"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
        If rsTmp.EOF Then
            stbInfo.Panels(1).Text = "�޲�����Ϣ"
        Else
            PatientDate = rsTmp(8)
            With stbInfo.Panels
                .Item(4).Text = IIf(PatientType = 0, "����ţ�" & rsTmp(0), "סԺ�ţ�" & rsTmp(1))
                .Item(1).Text = "������" & rsTmp(2) & "���Ա�" & rsTmp(3) & "�����䣺" & rsTmp(4)
                
                mstr�Ա� = rsTmp(3)
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
'    With Me.stbAdvInfo.Panels
'        .Item(1).Text = "��Ŀ��" + strDiagName
'        .Item(2).Text = "ҽ�����ݣ�" + strҽ������
'    End With
'    Me.stbDrAdviceInfo.Panels(1).Text = "ҽ�����У�" + strDrAdvice
    
    '�ж��ܷ�༭
    ProFile1(0).AllowEdit = False ' bAllowEdit
    ProFile1(1).AllowEdit = Not ReadOnly
    '����˵���������
    Me.mnuFileSave.Visible = Not ReadOnly: Me.mnuFileSplit(1).Visible = Not ReadOnly
    Me.tbrMain.Buttons("����").Visible = Not ReadOnly
    Me.mnuEdit_Clear.Visible = Not ReadOnly
    
    iTabIndex = -1
    TabFile.Tabs.Clear
    If alngFileID(0) > -1 And mblnShow���� Then TabFile.Tabs.Add , "����", "����(&S)", "����"
    TabFile.Tabs.Add , "����", "����(&B)", "����"
    TabFile.Tabs("����").Selected = True
    '����Tab
    Me.TabFile.Visible = True

    Set ParentForm = frmParent
    
    SetItemFormat
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
End Sub

Public Sub ShowMe_Request(frmParent As Object, ByVal lng����ID As Long, ByVal var��ҳ��Һ� As Variant, ByVal lng����ID As Long, _
    ByVal b��ʿվ As Boolean, Optional ByVal ModalWindow As Boolean = True, Optional ByVal lngǰ��ID As Long = 0)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '������Ŀ����
    Dim strDrAdvice As String 'ҽ������
    Dim bAllowEdit As Boolean
    Dim strSQL As String
    
    On Error Resume Next
    '��ʼ��
    mblnMoved = False
    
    alngFileID(0) = 0: PatientID = lng����ID: CheckID = CStr(var��ҳ��Һ�)
    PatientType = IIf(TypeName(var��ҳ��Һ�) = "String", 0, 1): FileTypeID = lng����ID: bSample = False: AdviceID = 0
    bln��ʿվ = b��ʿվ: mlngǰ��ID = lngǰ��ID
    
    Me.stbThis.Panels(3).Visible = False: Me.stbThis.Panels(4).Visible = False
        
    strSQL = "Select Count(*)" + _
        " From �����ļ���� Where �����ļ�ID=[1] And ��дʱ��=1"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
    If rsTmp(0) = 0 Then
        MsgBox "δ����������Ŀ�����ܱ༭", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    '��ʼ������
    
    '�ж��ܷ�༭����
    bAllowEdit = True
    
    iCurrElementIndex = 0

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
    
    With Me.lvwDemo.ColumnHeaders
        .Clear
        .Add , "����", "����", 1800
        .Add , "˵��", "˵��", 1800
    End With

    '��ȡ��ѡԪ���б�
'    GetElementList
'    mnuOrder_Add_FileList
    '��ȡ������Ϣ
    PatientDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    Me.Caption = "�����ļ�(����)"
    stbInfo.Visible = True
        
    strSQL = "Select ���� From �����ļ�Ŀ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, FileTypeID)
    If Not rsTmp.EOF Then Me.Caption = rsTmp(0) + "(����)"
    
    If PatientType = 0 Then
        strSQL = "Select Nvl(a.�����,0),Nvl(a.סԺ��,0),a.����,Nvl(a.�Ա�,' '),Nvl(a.����,' '),nvl(b.����,' ') As ����,nvl(c.����,' ') As ����,a.��ǰ����,a.����ʱ��,d.���˿���ID " + _
        "From ������Ϣ a,���ű� b,���ű� c,���˷��ü�¼ d Where a.����ID=[1] And a.��ǰ����ID=b.ID(+) And a.��ǰ����ID=c.ID(+) And " + _
        "d.��¼����=4 And d.��¼״̬ In (1,3) And d.���=1 And d.�����־=1 And d.����id=a.����id And d.��ʶ��=a.�����"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID)
    Else
        strSQL = "Select Nvl(a.�����,0),Nvl(a.סԺ��,0),a.����,Nvl(a.�Ա�,' '),Nvl(a.����,' '),nvl(b.����,' ') As ����,nvl(c.����,' ') As ����,a.��ǰ����,a.��Ժʱ��,d.��Ժ����ID " + _
        "From ������Ϣ a,���ű� b,���ű� c,������ҳ d Where a.����ID=[1] And a.��ǰ����ID=b.ID(+) And a.��ǰ����ID=c.ID(+) And " + _
        "d.��ҳID=[2] And d.����ID=a.����ID"
        Set rsTmp = OpenSQLRecord(strSQL, "zlCISCore", PatientID, CheckID)
    End If
    DeptID = UserInfo.����ID
    If rsTmp.EOF Then
        stbInfo.Panels(1).Text = "�޲�����Ϣ"
    Else
        PatientDate = rsTmp(8)
        lng���˿���ID = rsTmp(9)
        DeptID = rsTmp(9)
        With stbInfo.Panels
            .Item(4).Text = IIf(PatientType = 0, "����ţ�" & rsTmp(0), "סԺ�ţ�" & rsTmp(1))
            .Item(1).Text = "������" & rsTmp(2) & "���Ա�" & rsTmp(3) & "�����䣺" & rsTmp(4)
            
            mstr�Ա� = rsTmp(3)
            If PatientType = 0 Then
                .Item(2).Visible = False: .Item(3).Visible = False
            Else
                .Item(2).Text = "���ң�" & rsTmp(5)
                .Item(3).Text = "������" & rsTmp(6) & "�����ţ�" & NVL(rsTmp(7))
            End If
        End With
        
        Me.Caption = rsTmp(2) + "-" + Me.Caption
    End If
    
    ProFile1(0).AllowEdit = bAllowEdit
    '����˵���������
    iTabIndex = -1
    TabFile.Tabs.Clear
    TabFile.Tabs.Add , "����", "����(&S)", "����"
    TabFile.Tabs("����").Selected = True
    '����Tab
    Me.TabFile.Visible = False: Me.ProFile1(1).Visible = False

    '��ʼ������
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    '��ʼҽ���б�
    Call Get����ҽ��(CLng(PatientID), bln��ʿվ, "", 0, Me.cboҽ��, PatientType + 1)
    
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

    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
        "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ���� C Where A.ID=B.������ĿID And A.ID=C.������ĿID " + _
        "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN([1],3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And B.�����ļ�ID=[2] And Ӧ�ó���=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Name, PatientType + 1, FileTypeID)

    If rsTmp.EOF Then Exit Sub

    intType = -1
    If rsTmp!���ID = "D" And zlCommFun.NVL(GetItemField(rsTmp!������ĿID, "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intType = 0
    ElseIf rsTmp!���ID = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intType = 1
    ElseIf InStr(",7,8,", rsTmp!���ID) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intType = 2
    ElseIf rsTmp!���ID = "C" Then
        '������Ŀѡ�����걾
        intType = 4
    End If
    
    rsTmp.MoveFirst: If rsTmp.RecordCount = 1 Then ifInitItem = True '��Ϊֻ��һ����Ŀ����������ѡ�񣬽�������ʱֱ����ʾ������Ŀ

    SetItemFormat
End Sub

Private Sub SetItemFormat()   '����������Ŀ������ʾ��ʽ
    Select Case intType
        Case 0
            Me.lblҽ������.Caption = "�����Ŀ": Me.lbl����.Caption = "��鲿λ": Me.cmdExt.ToolTipText = "ѡ���鲿λ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����ʽ": Me.cmdExt.ToolTipText = "ѡ������ʽ"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
        Case 4
            Me.lblҽ������.Caption = "������Ŀ": Me.lbl����.Caption = "����걾": Me.cmdExt.ToolTipText = "ѡ�����걾"
            Me.lbl����.Visible = True: Me.txt����.Visible = True: Me.cmdExt.Visible = True
            Me.lbl�ɼ�.Visible = True: Me.txt�ɼ�.Visible = True: Me.cmd�ɼ�.Visible = True
        Case Else
            Me.lbl����.Visible = False: Me.txt����.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub EnableEditMenu(ByVal bAllowEdit As Boolean)
    Dim i As Integer
    Dim strSinglePriv As String
    
'    Me.mnuFileSave.Visible = bAllowEdit: Me.mnuFileSplit(1).Visible = bAllowEdit
    Me.mnuEdit.Visible = bAllowEdit
'    Me.mnuOrder_1.Visible = bAllowEdit
    For i = 1 To Me.tbrMain.Buttons.Count
        If Me.tbrMain.Buttons(i).Description = "�༭" Then Me.tbrMain.Buttons(i).Visible = bAllowEdit
    Next
    
    '�����ӡ����ˡ����ص�Ȩ��
    strSinglePriv = Left(mstrPrivs, 1)
    mnuEdit_Auditing.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("���").Visible = (strSinglePriv = 1)
    
    strSinglePriv = Mid(mstrPrivs, 2, 1)
    mnuEdit_Rollback.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("����").Visible = (strSinglePriv = 1)
    
    If mstrPrivs Like "00*" Then
        mnuEdit_0.Visible = False
        tbrMain.Buttons("Split_4").Visible = False
    End If
    
    strSinglePriv = Mid(mstrPrivs, 3, 1)
    mnuPreview.Visible = (strSinglePriv = 1)
    mnuPrint.Visible = (strSinglePriv = 1)
    tbrMain.Buttons("Ԥ��").Visible = (strSinglePriv = 1)
    tbrMain.Buttons("��ӡ").Visible = (strSinglePriv = 1)
    
    
'    '������˺Ͳ���Ȩ��
'    If mstrPrivs Like "00*" Then
'        mnuEdit_0.Visible = False
'        mnuEdit_Auditing.Visible = False
'        mnuEdit_Rollback.Visible = False
'
'        tbrMain.Buttons("Split_4").Visible = False
'        tbrMain.Buttons("���").Visible = False
'        tbrMain.Buttons("����").Visible = False
'    Else
'        mnuEdit_Auditing.Visible = (Mid(mstrPrivs, 1, 1) = 1)
'        mnuEdit_Rollback.Visible = (Mid(mstrPrivs, 2, 1) = 1)
'
'        tbrMain.Buttons("���").Visible = (Mid(mstrPrivs, 1, 1) = 1)
'        tbrMain.Buttons("����").Visible = (Mid(mstrPrivs, 2, 1) = 1)
'    End If
End Sub

Private Sub cmd�ɼ�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID))
    Me.txt�ɼ�.SetFocus
    If Not rsTmp Is Nothing Then
        Me.cmd�ɼ�.Tag = rsTmp("ID")
        Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub FileList_Click(Index As Integer)
    If MsgBox("���ز���ʾ���󣬵�ǰ�������ݽ������ǣ��Ƿ������", _
        vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    BeginShowProgress "���ڼ��أ�"
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
            BeginShowProgress "���ڼ������룺"
            ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, Me.prbRefresh, mlngǰ��ID, , , mblnMoved
            If alngFileID(0) = 0 Then Call ProFile1(0).SetDiagItem(lngClinicID, str�걾��λ)
        End If
        Me.MousePointer = vbHourglass
        BeginShowProgress "���ڼ��ر��棺"
        ProFile1(1).ShowFile IIf(alngFileID(1) = 0, "", CStr(alngFileID(1))), PatientID, CheckID, PatientType, FileTypeID, bSample, 2, Me.prbRefresh, mlngǰ��ID, AdviceID, SendNO, mblnMoved
        If alngFileID(1) = 0 Then Call ProFile1(1).SetDiagItem(lngClinicID, str�걾��λ)
        ProFile1(1).SetActiveElement 1
    Else
        If ProFile1(0).Tag = "" Then Exit Sub
        
        ProFile1(0).Tag = ""
        Me.MousePointer = vbHourglass
        BeginShowProgress "���ڼ������룺"
        ProFile1(0).ShowFile IIf(alngFileID(0) = 0, "", CStr(alngFileID(0))), PatientID, CheckID, PatientType, FileTypeID, bSample, 1, Me.prbRefresh, mlngǰ��ID, , , mblnMoved
        If alngFileID(0) = 0 Then Call ProFile1(0).SetDiagItem(lngClinicID, str�걾��λ)
        ProFile1(0).SetActiveElement 1
    End If
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
    
    If picAdvice.Enabled Then
        Me.txtҽ������.SetFocus
        If ifInitItem Then Call txtҽ������_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    '�й�ҽ���Ĳ���
    mstrLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    '��ʾ����ģ��
    mnuTemplate.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "����ģ��", "0"))
    mnuPatientInformation.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", "0"))
    tbrMain.Buttons("��ʾ").Value = IIf(mnuTemplate.Checked, tbrPressed, tbrUnpressed)
    Me.tvwElement.Visible = mnuTemplate.Checked
    
    'Ƥ�Խ����Чʱ��
    gint�����Ǽ���Ч���� = Val(GetSysParVal(2))
    
    '��һ�δ���Activate�¼�ʱҪ���ص���
    ProFile1(0).Tag = "Loading": ProFile1(1).Tag = "Loading"
    ProFile1(0).ifShowDiagItem = False: ProFile1(1).ifShowDiagItem = False
    
    '---------Ȩ�޿���-------------
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
    With Me.chk����
        .Left = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Width
        If .Left < Me.txt�ɼ�.Left + Me.txt�ɼ�.Width + lngDistance Then .Left = Me.txt�ɼ�.Left + Me.txt�ɼ�.Width + lngDistance
    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl��ʼʱ��.Left - Me.cmdSel.Width - Me.txtҽ������.Left - lngDistance - _
        Me.lbl����.Width - Me.cmdExt.Width - 60) / 2
    With Me.txtҽ������
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl����.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txtҽ������
        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl��ʼʱ��.Left - Me.txtƵ��.Left - Me.txtƵ��.Width - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30) - _
        (Me.lbl������λ.Width + Me.lbl����.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl����.Left = Me.txtƵ��.Left + Me.txtƵ��.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    Me.lbl����.Left = Me.lbl������λ.Left + Me.lbl������λ.Width + lngDistance
    With Me.txt����
        .Left = Me.lbl����.Left + Me.lbl����.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl������λ.Left = Me.txt����.Left + Me.txt����.Width + 30
    
    With Me.cboҽ��
        .Left = Me.txt����.Left
        .Width = picAdvice.Width - Me.lbl��ʼʱ��.Left - .Left
    End With
    Me.lbl����ҽ��.Left = Me.cboҽ��.Left - Me.lbl����ҽ��.Width
    
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
        
        If Not Me.TabFile.Visible Then  '����ʱ��ͬʱ��������ͱ���
            If MsgBox("�Ƿ񱣴��´������", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                mnuFileSave_Click
            End If
        Else
            If Val(GetSetting("ZLSOFT", "����ģ��\zl9Pacswork", "���Խ��������", 0)) = 1 Then
                If MsgBox("�Ƿ񱣴���д�ı���", vbQuestion + vbYesNo, gstrSysName) = vbYes Then SaveFile
            Else
                SaveFile
            End If
        End If
    End If
'    zlCommFun.OpenIme False
    
    Call SaveWinState(Me, App.ProductName)
    '������ʾģ��ѡ��
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "����ģ��", IIf(mnuTemplate.Checked, 1, 0))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", IIf(mnuPatientInformation.Checked, 1, 0))
    On Error Resume Next
    ParentForm.EditFile_UnLoad Me.hWnd  '�����ϼ����ڱ༭�ѹر�
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
        BeginShowProgress "����ˢ�£�"
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
'    If MsgBox("ȷ����˸ñ�����", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
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
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngAdviceID & "," & lngSendNO & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
        strSQL = "ZL_Ӱ����_STATE(" & lngAdviceID & "," & lngSendNO & ",5)"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        strSQL = "ZL_����ҽ��ִ��_Finish(" & lngAdviceID & "," & lngSendNO & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
        strSQL = "ZL_Ӱ����_STATE(" & lngAdviceID & "," & lngSendNO & ",6)"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "�������"
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
        strSQL = "Select a.ID,nvl(b.����,' ') From ���˲������� a,����Ԫ��Ŀ¼ b Where a.Ԫ�ر���=b.���� And a.ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngContentID)
        
        If Not rsTmp.EOF Then ProFile1(iTabIndex).CopyElement iCurrElementIndex, rsTmp("ID"), rsTmp(1)
    End If
End Sub

Private Sub mnuEdit_Exchange_Click()
    If MsgBox("���������ݽ��������ı������ݣ��Ƿ����", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    ProFile1(iTabIndex).ChangeToText iCurrElementIndex
    
    If Not Me.mnuEdit_Text.Checked Then
        mnuEdit_Text_Click
    Else
        If Not ProFile1(iTabIndex).ShowText(iCurrElementIndex, True) Then Me.mnuEdit_Text.Checked = False: Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuEdit_Map_Click()
    ProFile1(iTabIndex).EditElement iCurrElementIndex
End Sub

Private Sub mnuEdit_Rollback_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    If MsgBox("ȷ��Ҫ���ظñ�����", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
    On Error GoTo DBError
    strSQL = "Select Nvl(ִ�й���,0) As ִ�й��� From ����ҽ������ Where ҽ��ID=[1] And ���ͺ�=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, SendNO)
    If rsTmp.EOF Then Exit Sub
    
    If rsTmp(0) <> 6 Then
        strSQL = "ZL_Ӱ����_STATE(" & AdviceID & "," & SendNO & ",5)"
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
        MsgBox "�ñ����ı�û�����ݣ����ܴ�Ϊģ�塣", vbInformation, gstrSysName
        Exit Sub
    End If
    frmBillSave.ShowMe Me, ProFile1(iTabIndex).ElementID(iCurrElementIndex), ProFile1(iTabIndex).CurrentText(iCurrElementIndex)
End Sub

Private Sub mnuEdit_Text_Click()
    If ProFile1(iTabIndex).ShowText(iCurrElementIndex, Not Me.mnuEdit_Text.Checked) Then Me.mnuEdit_Text.Checked = Not Me.mnuEdit_Text.Checked
    Me.tbrMain.Buttons("�ı�").Value = IIf(Me.mnuEdit_Text.Checked, tbrPressed, tbrUnpressed)
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
    If Me.TabFile.Visible Then  '����ʱ��ͬʱ��������ͱ���
        If Val(GetSetting("ZLSOFT", "����ģ��\zl9Pacswork", "���Խ��������", 0)) = 0 Then
            iMsgReturn = MsgBox("��ȷ�ϱ������Ƿ�Ϊ���ԣ�" & vbCrLf & "ѡ��ȡ����������档", vbYesNoCancel + vbQuestion + vbDefaultButton1, gstrSysName)
            If iMsgReturn = vbCancel Then Exit Function
            iMsgReturn = IIf(iMsgReturn = vbYes, 1, 0)
        Else
            iMsgReturn = 0
        End If
        
        If alngFileID(0) > -1 And alngFileID(1) = 0 Then 'Ҫ��������
            
            If mblnShow���� Then
'                If MsgBox("���汣��ʱ��ͬʱ�������룬֮�����뽫�����޸ģ��Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            End If
            
            '��������
            sTmpFileID = ProFile1(0).SaveFile
            If Len(sTmpFileID) > 0 Then
                alngFileID(0) = CLng(sTmpFileID)
                
                CommitData 0
            Else
                Exit Function
            End If
        End If
        '���汨��
        sTmpFileID = ProFile1(1).SaveFile
        If Len(sTmpFileID) > 0 Then
            alngFileID(1) = CLng(sTmpFileID)
            
            CommitData 1, iMsgReturn
            
            ProFile1(0).AllowEdit = False '�������ٱ༭����
            SaveFile = True: Exit Function
        Else
            Exit Function
        End If
    Else
        '��������
        
        If Me.picAdvice.Enabled Then
            If MsgBox("���뱣���ϵͳ���Զ�������ʱҽ����" + Chr(13) + "������Ŀ�������޸ģ��Ƿ�Ҫ���棿", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
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
'��дҽ���������
Private Sub CommitData(ByVal iCommitType As Integer, Optional ByVal iCheckResult As Integer = -1)
    On Error GoTo DBError
    If iCommitType = 1 Then '����
        If iCheckResult = -1 Then
            gcnOracle.Execute "ZL_���Ƶ���_����('" & sCheckNo & "'," & iRecordType & "," & alngFileID(1) & "," & _
                IIf(blnҽ��ִ��, 1, 0) & "," & AdviceID & ")", , adCmdStoredProc
        Else
            gcnOracle.Execute "ZL_���Ƶ���_����('" & sCheckNo & "'," & iRecordType & "," & alngFileID(1) & "," & _
                IIf(blnҽ��ִ��, 1, 0) & "," & AdviceID & "," & iCheckResult & ")", , adCmdStoredProc
        End If
    Else '����
        gcnOracle.Execute "ZL_���Ƶ���_����(" & AdviceID & "," & alngFileID(0) & ")", , adCmdStoredProc
    End If
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub
'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If Len(Trim(strAdviceText)) = 0 Then
        ValidAdvice = False
        MsgBox "��������������Ŀ��", vbInformation, gstrSysName
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Len(Trim(strSequence)) = 0 Then
        ValidAdvice = False
        MsgBox "����ָ��Ƶ�ʣ�", vbInformation, gstrSysName
        Me.txtƵ��.SetFocus: Exit Function
    End If
    If Not Check��ʼʱ��(CStr(Me.txt��ʼʱ��)) Then
        ValidAdvice = False
        Me.txt��ʼʱ��.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt����)) = 0 Then
        ValidAdvice = False
        MsgBox "������������", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt����)) = 0 And Me.txt����.Enabled Then
        ValidAdvice = False
        MsgBox "�����뵥����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
    If Val(Me.txt����) > Val(Me.txt����) Then
        ValidAdvice = False
        MsgBox "�������ܴ���������", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Function
    End If
End Function
'����ҽ��
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
    Dim lng��������ID As Long, strDoctor As String, i As Integer
    Dim strִ�п���ID As String, strִ�п���ID1 As String
    Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer, tmpintִ������ As Integer
    Dim rsDept As ADODB.Recordset

    gcnOracle.BeginTrans
    On Error GoTo DBError
    
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    strSQL = "Select Max(���) From ����ҽ����¼ Where ����ID=[1]" & _
        " And " & IIf(PatientType = 1, "��ҳID=[2]", "�Һŵ�=[2]")
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, CheckID)
    If IsNull(rsTmp(0)) Then
        iMaxSeq = 0
    Else
        iMaxSeq = rsTmp(0)
    End If
    
    lng��������ID = Get��������ID(Me.cboҽ��.ItemData(Me.cboҽ��.ListIndex), lng���˿���ID, PatientType + 1)
    i = InStr(Me.cboҽ��.Text, "-")
    If i > 0 Then strDoctor = Mid(Me.cboҽ��, i + 1)
    If Len(Me.cboִ�п���.Text) = 0 Then
        strִ�п���ID = "NULL"
    Else
        strִ�п���ID = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex)
    End If
    
    tmpstr��� = str���: tmplngClinicID = lngClinicID: tmpint�Ƽ����� = int�Ƽ�����
    tmpintִ������ = intִ������
    If intType = 4 Then
        '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
        strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
        If rsTmp.State = adStateOpen Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Me.cmd�ɼ�.Tag)
        tmpstr��� = rsTmp("���"): tmplngClinicID = rsTmp("ID"): tmpint�Ƽ����� = NVL(rsTmp("�Ƽ�����"), 0)
        tmpintִ������ = NVL(rsTmp("ִ�п���"), 0)
        'ȡ�ɼ���ʽ��ִ�в���
        Set rsDept = GetExeDepart(rsTmp("ID"), PatientType + 1, DeptID)
        If rsDept Is Nothing Then
            strִ�п���ID1 = "NULL"
        Else
            strִ�п���ID1 = rsDept("ID")
        End If
    End If
    
    If intType <> 4 Then
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr��� & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
            "'" & str�걾��λ & "','" & strSequence & "'," & _
            IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
            IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
            "'" & str�����λ & "',NULL," & _
            tmpint�Ƽ����� & "," & _
            strִ�п���ID & "," & _
            tmpintִ������ & "," & Me.chk����.Value & "," & _
            IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
            "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
            IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If
    '�������ҽ��
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            iMaxSeq = iMaxSeq + 1
            With rsRelativeAdvice
                strSQL = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
                    "0,1," & _
                    "1,'" & .Fields("���") & "'," & _
                    .Fields("ID") & ",NULL,NULL," & _
                    IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
                    IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
                    "'" & Replace(.Fields("����"), "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
                    "'" & IIf(intType = 4, str�걾��λ, .Fields("�걾��λ")) & "','" & strSequence & "'," & _
                    IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
                    IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
                    "'" & str�����λ & "',NULL," & _
                    .Fields("�Ƽ�����") & "," & _
                    strִ�п���ID & "," & _
                    .Fields("ִ�п���") & "," & Me.chk����.Value & "," & _
                    IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
                    "NULL," & _
                    lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
                    "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
                    IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
                gcnOracle.Execute strSQL, , adCmdStoredProc
                
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    
    If intType = 4 Then
        '��������Ĳɼ���ʽ�ŵ����
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & (PatientType + 1) & "," & PatientID & "," & IIf(PatientType = 1, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr��� & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            IIf(Len(Trim(Me.txt����)) = 0, "NULL", Me.txt����) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txtҽ������, "'", "''") & "'," & _
            "'" & str�걾��λ & "','" & strSequence & "'," & _
            IIf(lngƵ�ʴ��� = 0, "NULL", lngƵ�ʴ���) & "," & _
            IIf(lngƵ�ʼ�� = 0, "NULL", lngƵ�ʼ��) & "," & _
            "'" & str�����λ & "',NULL," & _
            tmpint�Ƽ����� & "," & _
            strִ�п���ID1 & "," & _
            tmpintִ������ & "," & Me.chk����.Value & "," & _
            IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "NULL,", "To_Date('" & Format(Me.txt��ʼʱ��.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng���˿���ID & "," & lng��������ID & ",'" & strDoctor & "'," & _
            "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & IIf(PatientType = 1, "", CheckID) & "'," & _
            IIf(mlngǰ��ID = 0, "Null", mlngǰ��ID) & ")"
        gcnOracle.Execute strSQL, , adCmdStoredProc
    End If

    gcnOracle.CommitTrans
    AdviceID = lngAdviceID
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "����ҽ������"
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
    BeginShowProgress "����ˢ�£�"
    ProFile1(iTabIndex).DeleteElement iCurrElementIndex, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault

    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuOrder_Demo_Click()
    tbrMain_ButtonClick tbrMain.Buttons("Ԫ��")
End Sub

Private Sub mnuOrder_Insert_Click()
    tbrMain_ButtonClick tbrMain.Buttons("����")
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
        If MsgBox("�ò����������ģ���ӡ֮ǰϵͳ������÷ݲ������Ƿ����", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1(iTabIndex).Modified Then _
            If MsgBox("��ӡ֮ǰ�Ƿ񱣴�÷ݲ���", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then If Not SaveFile Then Exit Sub
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
        '��ӡ����
        PrintDiagReport AdviceID, SendNO, Me, 1, Me.picBuffer, mblnMoved
    End If
End Sub

Private Sub mnuPrint_Click()
    Dim rsTmp As New ADODB.Recordset
    
    Dim intPage As Integer
    
    If alngFileID(iTabIndex) = 0 Then
        If MsgBox("�ò����������ģ���ӡ֮ǰϵͳ������÷ݲ������Ƿ����", vbDefaultButton1 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        Else
            If Not SaveFile Then Exit Sub
        End If
    Else
        If ProFile1(iTabIndex).Modified Then _
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
    If iTabIndex = 0 Then
        If bSample Then
            PrintOutCase Me, Printer, 0, True, 1, 0, alngFileID(iTabIndex), False, 0, 1
        Else
            PrintOutCase Me, Printer, 5, True, -1 * CLng(Val(alngFileID(iTabIndex))), CLng(PatientID), CheckID, False, 0, 1
        End If
    Else
        '��ӡ����
        PrintDiagReport AdviceID, SendNO, Me, 2, Me.picBuffer, mblnMoved
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
    BeginShowProgress "���ڼ��أ�"
    If iTabIndex = 0 Then
        ProFile1(iTabIndex).ShowFile IIf(alngFileID(iTabIndex) = 0, "", CStr(alngFileID(iTabIndex))), PatientID, CheckID, PatientType, FileTypeID, bSample, iTabIndex + 1, Me.prbRefresh, mlngǰ��ID, , , mblnMoved
    Else
        ProFile1(iTabIndex).ShowFile IIf(alngFileID(iTabIndex) = 0, "", CStr(alngFileID(iTabIndex))), PatientID, CheckID, PatientType, FileTypeID, bSample, iTabIndex + 1, Me.prbRefresh, mlngǰ��ID, AdviceID, SendNO, mblnMoved
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
    tbrMain.Buttons("��ʾ").Value = IIf(mnuTemplate.Checked, tbrPressed, tbrUnpressed)
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
        Case "����"
            If iTabIndex = 0 Then Exit Sub
            
            iTabIndex = 0
        Case "����"
            If iTabIndex = 1 Then Exit Sub
            
            iTabIndex = 1
    End Select
            
    Me.ProFile1(0).Visible = False
    Me.ProFile1(1).Visible = False
    picFile_Resize
    '���ñ༭�˵�
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
            Me.PopupMenu Me.mnuOrder_Add
        Case "Ԫ��"
            With Me.lvwDemo
                GetElementDemoList ProFile1(iTabIndex).ElementID(iCurrElementIndex)
                .Left = Button.Left
                .Top = Button.Top + Button.Height + 30
                .ZOrder 0: .Visible = True: lvwItem.Visible = False
                .SetFocus
            End With
        Case "ɾ��"
            mnuOrder_Delete_Click
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
        Case "���"
            mnuEdit_Auditing_Click
        Case "����"
            mnuEdit_Rollback_Click
        Case "��ʾ"
            mnuTemplate_Click
        Case "ģ��"
            mnuEdit_Template_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
            mnuExit_Click
    End Select
End Sub

Private Sub ShowEditMenu(ElementType As Integer)
    If Not ProFile1(iTabIndex).AllowEdit Then Exit Sub
    Select Case ElementType
        Case 2 '������
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = True
            Me.tbrMain.Buttons("�ı�").Value = IIf(ProFile1(iTabIndex).IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("ת��").Enabled = True
            Me.tbrMain.Buttons("����").Enabled = True
            Me.tbrMain.Buttons("�༭").Enabled = False
            Me.tbrMain.Buttons("ģ��").Enabled = False
        Case 3 '���ͼ
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = False
            Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
            Me.tbrMain.Buttons("ת��").Enabled = False
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�༭").Enabled = True
            Me.tbrMain.Buttons("ģ��").Enabled = False
        Case 4 'ר��ֽ
            Me.tbrMain.Buttons("����").Enabled = False
            Me.tbrMain.Buttons("�ı�").Enabled = True
            Me.tbrMain.Buttons("�ı�").Value = IIf(ProFile1(iTabIndex).IsText(iCurrElementIndex), tbrPressed, tbrUnpressed)
            Me.tbrMain.Buttons("ת��").Enabled = True
            Me.tbrMain.Buttons("����").Enabled = True
            Me.tbrMain.Buttons("�༭").Enabled = False
            Me.tbrMain.Buttons("ģ��").Enabled = False
        Case Else
            Me.tbrMain.Buttons("����").Enabled = IIf(ElementType = 0, True, False)
            Me.tbrMain.Buttons("�ı�").Enabled = False
            Me.tbrMain.Buttons("�ı�").Value = tbrUnpressed
            Me.tbrMain.Buttons("ת��").Enabled = False
            Me.tbrMain.Buttons("����").Enabled = True
            Me.tbrMain.Buttons("�༭").Enabled = False
            Me.tbrMain.Buttons("ģ��").Enabled = IIf(ElementType = 0, True, False)
    End Select
    
    Me.mnuEdit_Copy.Enabled = Me.tbrMain.Buttons("����").Enabled
    Me.mnuEdit_Char.Enabled = Me.tbrMain.Buttons("����").Enabled
    Me.mnuEdit_Map.Enabled = Me.tbrMain.Buttons("�༭").Enabled
    Me.mnuEdit_Text.Enabled = Me.tbrMain.Buttons("�ı�").Enabled
    Me.mnuEdit_Text.Checked = IIf(Me.tbrMain.Buttons("�ı�").Value = tbrPressed, True, False)
    Me.mnuEdit_Exchange.Enabled = Me.tbrMain.Buttons("ת��").Enabled
    Me.mnuEdit_Template.Enabled = Me.tbrMain.Buttons("ģ��").Enabled
    
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
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuOrder_Add_FileList()
    Dim rsFileList As New ADODB.Recordset
    Dim i As Integer, iNum As Integer
    Dim strSQL As String
    
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
            " ����ʾ��Ŀ¼ Where ID=" & alngFileID(iTabIndex), Me.Caption
            
            FileTypeID = rsFileList(0)
        Else
            zlDatabase.OpenRecordset rsFileList, "Select �ļ�ID From" + _
            " ���˲�����¼ Where ID=" & alngFileID(iTabIndex), Me.Caption
            
            FileTypeID = rsFileList(0)
        End If
    End If
    
    strSQL = "Select a.ID,a.���� From ����ʾ��Ŀ¼ a" + _
        " Where a.�ļ�ID=[1] And a.����=1" + _
        IIf(bSample, " And a.ID<>[2]", "") + _
        IIf(bSample, "", " And (a.����ID=[3] Or" + _
        " a.����ID Is Null)")
    Set rsFileList = OpenSQLRecord(strSQL, Me.Caption, FileTypeID, alngFileID(iTabIndex), UserInfo.����ID)
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
    Dim strSQL As String
    
    Me.lvwDemo.ListItems.Clear
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select a.ID,a.����,a.˵�� From ����ʾ��Ŀ¼ a" + _
        " Where a.Ԫ��ID=[1] And a.����=2" + _
        IIf(bSample, "", " And (a.����ID=[2] Or" + _
        " a.����ID Is Null)")
    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, ElementID, UserInfo.����ID)
    If rsTemp.EOF Then Exit Sub
    With rsTemp
        Me.lvwDemo.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwDemo.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Ԫ��": objItem.SmallIcon = "Ԫ��"
            objItem.SubItems(Me.lvwDemo.ColumnHeaders("˵��").Index - 1) = IIf(IsNull(!˵��), "", !˵��)
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

'========������ҽ���༭==========

Private Sub cboִ�п���_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk����_Click()
    On Error Resume Next
    Me.txtҽ������.SetFocus
End Sub

Private Sub chk����_GotFocus()
    EnableEditMenu False
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk��ʼʱ��_Click()
    On Error Resume Next
    If Me.chk��ʼʱ��.Value = 1 Then
        Me.txt��ʼʱ��.Enabled = True: Me.txt��ʼʱ��.SetFocus
    Else
        Me.txt��ʼʱ��.Enabled = False
    End If
    
    If str��� = "D" Then
        strAdviceText = Get�����������(1, strClinicName)
    ElseIf str��� = "F" Then
        strAdviceText = Get�����������(2, strClinicName)
    End If
End Sub

Private Sub chk��ʼʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboҽ��_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ProFile1(iTabIndex).SetFocus
End Sub

Private Sub cmdExt_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim tmpExtData As String
    
    frmAdviceEditEx.mlngHwnd = Me.cboҽ��.hWnd 'txt����.Hwnd
    frmAdviceEditEx.mintType = IIf(intType = 4, 3, intType)
    frmAdviceEditEx.mint��Ч = 1
    frmAdviceEditEx.mstr�Ա� = mstr�Ա�
    If intType = 4 Then
        '������Ŀ
        frmAdviceEditEx.mlng��ĿID = 0 'Split(strExtData, ";")(0)
        frmAdviceEditEx.mstrExtData = strExtData ' Split(strExtData, ";")(1)
    Else
        frmAdviceEditEx.mlng��ĿID = lngClinicID
        frmAdviceEditEx.mstrExtData = strExtData
    End If
    frmAdviceEditEx.mint������� = PatientType + 1

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me

    If Not frmAdviceEditEx.mblnOK Then
        zlControl.TxtSelAll Me.txt����
        Me.txt����.SetFocus
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
        Case 0 '�����ϲ�λ
            Call AdviceSet�������(1, strExtData)
            strAdviceText = Get�����������(1, strClinicName)
            Me.txt���� = Get��λ����
        Case 1 '������Ŀ
            Call AdviceSet�������(2, strExtData)
            txtҽ������.Text = Get�����������(2, strClinicName)
            strAdviceText = Get�����������(2, strClinicName)
            Me.txt���� = Get��������
        Case 4 '������Ŀ
            strAdviceText = strClinicName & "(" & tmpExtData & ")"
            Me.txt���� = tmpExtData: str�걾��λ = tmpExtData
    End Select
    txt����.Tag = txt����.Text
    Me.txt����.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdExt_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset
    
    If intType = 4 Then
        '������Ŀ
        If LabsInput Then
            txtҽ������.Tag = txtҽ������.Text
            txt����.Tag = txt����.Text
            Me.txtҽ������.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            txt����.Text = txt����.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus
        End If
        Exit Sub
    End If
    
    With txtҽ������
        .Text = ""
        Set rsTmp = SelectDiagItem()
    End With
    
    If rsTmp Is Nothing Then 'ȡ����������
        '�ָ�ԭֵ
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus: Exit Sub
    End If
    '����Ŀ��¼��
    
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If AdviceInput(rsTmp) Then
        '��ʾ��ȱʡ���õ�ֵ
        txtҽ������.Tag = txtҽ������.Text
        txt����.Tag = txt����.Text
        Me.txtҽ������.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '�ָ�ԭֵ
        txtҽ������.Text = txtҽ������.Tag
        txt����.Text = txt����.Tag
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus
    End If
End Sub

Private Sub cmdSel_GotFocus()
    EnableEditMenu False
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
        
    int��Χ = 1
    strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
        " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,nvl(A.�����λ,' ') As �����λ" & _
        " From ����Ƶ����Ŀ A Where A.���÷�Χ=" & int��Χ & _
        " Order by A.����"
    vRect = GetControlRect(txtƵ��.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Ƶ��", , , , , , True, vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
        End If
        txtƵ��.Text = strSequence
        Call zlControl.TxtSelAll(txtƵ��)
        txtƵ��.SetFocus: Exit Sub
    End If
    Me.cmdƵ��.Tag = rsTmp("����"): Me.txtƵ�� = Me.cmdƵ��.Tag: strSequence = Me.cmdƵ��.Tag
    lngƵ�ʴ��� = rsTmp("Ƶ�ʴ���"): lngƵ�ʼ�� = rsTmp("Ƶ�ʼ��"): str�����λ = Trim(rsTmp("�����λ"))

    txtƵ��.SetFocus
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdƵ��_GotFocus()
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

Private Sub txt�ɼ�_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt�ɼ�)
End Sub

Private Sub txt�ɼ�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt�ɼ�.Text = txt�ɼ�.Tag Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID), Me.txt�ɼ�)
    If Not rsTmp Is Nothing Then
        Me.cmd�ɼ�.Tag = rsTmp("ID")
        Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�ɼ�_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txt�ɼ�.Text <> txt�ɼ�.Tag Then
        txt�ɼ�.Text = txt�ɼ�.Tag
    End If
End Sub

Private Sub txt����_GotFocus()
    EnableEditMenu False
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt����) Then Me.txt���� = 1: Exit Sub
    Me.txt���� = CInt(Me.txt����)
    If CInt(Me.txt����) < 1 Then Me.txt���� = 1
End Sub

Private Sub txt����_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt����_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt����)
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text = txt����.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        cmdExt_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txt����.Text <> txt����.Tag Then
        txt����.Text = txt����.Tag
    End If
End Sub

Private Sub txt��ʼʱ��_GotFocus()
    EnableEditMenu False
End Sub

Private Sub txt��ʼʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��ʼʱ��_Validate(Cancel As Boolean)
    On Error Resume Next
    If Not Check��ʼʱ��(CStr(txt��ʼʱ��)) Then
        Cancel = True
        txt��ʼʱ��.SetFocus
    Else
        If str��� = "D" Then
            strAdviceText = Get�����������(1, strClinicName)
        ElseIf str��� = "F" Then
            strAdviceText = Get�����������(2, strClinicName)
        End If
    End If
End Sub

Private Sub txtƵ��_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txtƵ��)
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdƵ��.Tag <> "" And txtƵ��.Text = strSequence And txtƵ��.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txtƵ��.Text = "" Then
            If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
        Else
            int��Χ = 1 '��ѡƵ��
            strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
                " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
                " From ����Ƶ����Ŀ A Where A.���÷�Χ=" & int��Χ & _
                " And (A.���� Like '" & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.����) Like '" & mstrLike & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.����) Like '" & mstrLike & UCase(txtƵ��.Text) & "%'" & _
                " Or Upper(A.Ӣ������) Like '" & mstrLike & UCase(txtƵ��.Text) & "%')" & _
                " Order by A.����"
            vRect = GetControlRect(txtƵ��.hWnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Ƶ��", , , , , , True, vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                End If
                txtƵ��.Text = strSequence
                Call zlControl.TxtSelAll(txtƵ��)
                txtƵ��.SetFocus: Exit Sub
            End If
            Me.cmdƵ��.Tag = rsTmp("����"): Me.txtƵ�� = Me.cmdƵ��.Tag: strSequence = Me.cmdƵ��.Tag
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtƵ��_Validate(Cancel As Boolean)
    If cmdƵ��.Tag <> "" And txtƵ��.Text <> strSequence Then
        txtƵ��.Text = strSequence
    End If
End Sub

Private Sub txtҽ������_GotFocus()
    EnableEditMenu False
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    On Error Resume Next
    If zlCommFun.ActualLen(txtҽ������.Text) > txtҽ������.MaxLength Then
        MsgBox "�������ݲ������� " & txtҽ������.MaxLength \ 2 & " �����ֻ� " & txtҽ������.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtҽ������.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txtҽ������_DblClick()
    If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
End Sub

Private Sub txtҽ������_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txtҽ������)
    End If
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtҽ������.Text = "" Then cmdSel_Click: Exit Sub
        If txtҽ������.Text = txtҽ������.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        With txtҽ������
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        If AdviceInput(rsTmp) Then
            '��ʾ��ȱʡ���õ�ֵ
            txtҽ������.Tag = txtҽ������.Text
            txt����.Tag = txt����.Text
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            txt����.Text = txt����.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txtҽ������.Text <> txtҽ������.Tag Then
        txtҽ������.Text = txtҽ������.Tag
    End If
End Sub

Private Sub txt����_GotFocus()
    EnableEditMenu False
    Call zlControl.TxtSelAll(Me.txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt����) Then Me.txt���� = 1: Exit Sub
    Me.txt���� = CInt(Me.txt����)
    If CInt(Me.txt����) < 1 Then Me.txt���� = 1
End Sub

'�ж��Ƿ�Ϊ�༭��
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function Check��ʼʱ��(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵���Ժʱ��
'2.��ʼʱ�����С����ֹʱ��
'3.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
'4.��¼��ҽ����ʼʱ�䲻�ܴ��ڵ�ǰʱ��
    Dim strInDate As String
    
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
        
    strInDate = Format(PatientDate, "yyyy-MM-dd HH:mm")
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵�" & IIf(PatientType = 0, "����", "��Ժ") & "ʱ�� " & strInDate & " ��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
'    If IsDate(strEnd) Then
'        If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(strEnd, "yyyy-MM-dd HH:mm") Then
'            strMsg = "ҽ���Ŀ�ʼִ��ʱ�����С��ִ����ֹʱ�䡣"
'            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    If DateDiff("n", CDate(strStart), zlDatabase.Currentdate) > 30 Then
        strMsg = "��ʼִ��ʱ�䲻��̫���ڵ�ǰʱ�䡣"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check��ʼʱ�� = True
End Function

Private Function SelectDiagItem() As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
        "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ���� C Where A.ID=B.������ĿID And A.ID=C.������ĿID " + _
        "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN(" & (PatientType + 1) & ",3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And (A.���� Like '" + txtҽ������ + "%' Or Upper(A.����) Like '" + mstrLike + txtҽ������ + "%' Or Upper(C.����) Like '" + mstrLike + UCase(txtҽ������) + "%') And B.�����ļ�ID=" & FileTypeID & " And Ӧ�ó���=" & (PatientType + 1)
            
    With txtҽ������
        Me.stbThis.Panels(2).Text = "��ѡ��������Ŀ..."
        Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "ѡ��������Ŀ", True, .Text, "", True, True, True, .Left + Me.picAdvice.Left + Me.Left, .Top + Me.picAdvice.Top + Me.Top, .Height, False, True)
        Me.stbThis.Panels(2).Text = ""
    End With
End Function

Private Function SelectCap(Optional ByVal lngItemID As Long = 0, Optional ByVal QryStr As String = "", Optional blnNotSelect As Boolean = False) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
    If Len(QryStr) > 0 Then
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A,������Ŀ���� C,�����÷����� D Where A.ID=C.������ĿID And A.ID=D.�÷�ID" + _
            " And A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & (PatientType + 1) & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
            " And D.��ĿID=" & lngItemID & _
            " And (A.���� Like '" + QryStr + "%' Or Upper(A.����) Like '" + mstrLike + QryStr + "%' Or Upper(C.����) Like '" + mstrLike + UCase(QryStr) + "%')"
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                "From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID" + _
                " And A.���='E' And A.��������='6'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                " And A.������� IN(" & (PatientType + 1) & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
                " And (A.���� Like '" + QryStr + "%' Or Upper(A.����) Like '" + mstrLike + QryStr + "%' Or Upper(C.����) Like '" + mstrLike + UCase(QryStr) + "%')"
        End If
    Else
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
            " And A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & (PatientType + 1) & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
            " And D.��ĿID=" & lngItemID
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                "From ������ĿĿ¼ A Where " + _
                " A.���='E' And A.��������='6'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                " And A.������� IN(" & (PatientType + 1) & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
        End If
    End If
    If blnNotSelect Then
        If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
        OpenRecord rsTmp, strSQL, Me.Caption
        If Not rsTmp.EOF Then Set SelectCap = rsTmp
    Else
        tmpRect = GetControlRect(Me.txt�ɼ�.hWnd)
        Set SelectCap = zlDatabase.ShowSelect(Me, strSQL, 0, "�ɼ���ʽ", True, , , , , True, _
            tmpRect.Left, tmpRect.Top, Me.txt�ɼ�.Height, , , True)
    End If
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String

    On Error GoTo errH

    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    txtҽ������.Text = rsInput!���� '��ʱ��ʾ

    '��Ҫ����������ݵ�һЩ��Ŀ
    '---------------------------------------------------------------------------------------------------------------
    intTmpType = -1
    If rsInput!���ID = "D" And zlCommFun.NVL(GetItemField(rsInput!������ĿID, "�����Ŀ"), 0) = 1 Then
        '��������Ŀ
        intTmpType = 0
        strHelpText = "��鲿λ"
    ElseIf rsInput!���ID = "F" Then
        '��������Ҫ����������Ŀ������ѡ�񸽼�����
        intTmpType = 1
        strHelpText = "��������������ʽ"
    ElseIf InStr(",7,8,", rsInput!���ID) > 0 Then
        '��ҩ�䷽(��ζ��ҩ���䷽����)
        intTmpType = 2
    ElseIf rsInput!���ID = "C" Then
        '������Ŀѡ�����걾
        intTmpType = 4
        strHelpText = "������Ŀ"
    End If

    If intTmpType <> -1 Then
        frmAdviceEditEx.mlngHwnd = Me.cboִ�п���.hWnd ' txtҽ������.Hwnd
        frmAdviceEditEx.mintType = intTmpType
        frmAdviceEditEx.mint��Ч = 1
        frmAdviceEditEx.mstr�Ա� = mstr�Ա�
        frmAdviceEditEx.mlng��ĿID = IIf(intTmpType = 4, FileTypeID, rsInput!������ĿID)
        frmAdviceEditEx.mstrExtData = IIf(intTmpType = 4, rsInput!������ĿID & ";" & NVL(rsInput("�걾��λ")), "") '��������Ŀ
        frmAdviceEditEx.mint������� = PatientType + 1

        On Error Resume Next
        Me.stbThis.Panels(2).Text = "��ѡ��" + strHelpText + "..."
        frmAdviceEditEx.Show 1, Me
        Me.stbThis.Panels(2).Text = ""
        On Error GoTo errH

        If Not frmAdviceEditEx.mblnOK Then Exit Function
        If frmAdviceEditEx.mstrExtData = "" Or (Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" And rsInput!���ID <> "F") Then Exit Function
        
        If rsInput!���ID = "D" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '������ϲ�λ��
            Call AdviceSet�������(1, strExtData)
            txtҽ������.Text = Get�����������(1, rsInput!����)
            strAdviceText = Get�����������(1, rsInput!����)
            Me.txt���� = Get��λ����
        ElseIf rsInput!���ID = "F" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '�����ĸ���������������Ŀ��
            Call AdviceSet�������(2, strExtData)
            txtҽ������.Text = Get�����������(2, rsInput!����)
            strAdviceText = Get�����������(2, rsInput!����)
            Me.txt���� = Get��������
        ElseIf rsInput!���ID = "C" And frmAdviceEditEx.mstrExtData <> "" Then
            '��ȡ�ɼ���ʽ
            Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
            If rsTmp Is Nothing Then
                MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
                Exit Function
            End If
            Me.cmd�ɼ�.Tag = rsTmp("ID")
            Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
            
            strAdviceText = txtҽ������.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str�걾��λ = Trim(rsInput("�걾��λ"))
            
            '������Ŀ
            strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
                "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
                "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
                "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
                "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ���� C Where A.ID=B.������ĿID And A.ID=C.������ĿID " + _
                "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                "And A.������� IN([1],3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
                IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
                " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
                " And A.ID=[2] And B.�����ļ�ID=[3] And Ӧ�ó���=[1]"
            If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
            Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType + 1, Split(Split(strExtData, ";")(0), ",")(0), FileTypeID)
            
            Call AdviceSet�������(3, strExtData)
            txtҽ������.Text = Get�����������(2, "")
            strAdviceText = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
            Me.txt���� = Split(strExtData, ";")(1)
            str�걾��λ = Me.txt����
        End If
    Else
        str�걾��λ = Trim(rsInput("�걾��λ"))
        txtҽ������.Text = txtҽ������.Text & "(" & str�걾��λ & ")"
        strAdviceText = txtҽ������.Text
        
        '������ϲ�λ��
        Call AdviceSet�������(1, "")
    End If
    
    '��ʼʱ��
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("ִ�а���ID") = 1 Then
        Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ��.Enabled = False
    Else
        Me.lbl��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Visible = False
        Me.txt��ʼʱ��.Enabled = True
    End If
    
    '����Ƶ��
    If rsInput("ִ��Ƶ��ID") = 1 Then
        Me.txtƵ��.Enabled = False: Me.txtƵ�� = "һ����": Me.cmdƵ��.Enabled = False
    Else
        Me.txtƵ��.Enabled = True: Me.txtƵ�� = "": Me.cmdƵ��.Enabled = True
    End If
    strSequence = Me.txtƵ��
    
    '����
    Me.txt���� = "1": Me.lbl������λ.Caption = rsInput("���㵥λ")
    
    '����
    If (rsInput("ִ��Ƶ��ID") = 0 And InStr(",1,2,", rsInput("���㷽ʽID")) > 0) _
                    Or InStr(",5,6,", rsInput("���ID")) > 0 Then
        Me.txt����.Enabled = True: Me.txt���� = "": Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = rsInput("���㵥λ")
    Else
        Me.txt����.Enabled = False: Me.txt���� = "": Me.txt����.BackColor = Me.BackColor: Me.lbl������λ.Caption = "" ' rsInput("���㵥λ")
    End If
    
    'ִ�п���
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType + 1, DeptID)
    If rsTmp Is Nothing Then
        Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cboִ�п���.Clear
        Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(0) = rsTmp("ID"): Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    Else
        Me.cboִ�п���.Clear
        Do While Not rsTmp.EOF
            Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(Me.cboִ�п���.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = True: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    End If
    
    '����ҽ��
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '����������Ŀ������ʾ��ʽ
    
    str��� = rsInput("���ID"): lngClinicID = rsInput("������ĿID"): Call ProFile1(0).SetDiagItem(lngClinicID, str�걾��λ)
    int�Ƽ����� = rsInput("�Ƽ�����ID"): intִ������ = rsInput("ִ�п���ID"): strClinicName = IIf(intType = 4, Me.txtҽ������, rsInput("����"))
    
    AdviceInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LabsInput() As Boolean
'���ܣ��༭������Ŀ
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String, rsInput As New ADODB.Recordset

    On Error GoTo errH
    
    intTmpType = 4
    strHelpText = "������Ŀ"

    frmAdviceEditEx.mlngHwnd = Me.cboִ�п���.hWnd ' txtҽ������.Hwnd
    frmAdviceEditEx.mintType = intTmpType
    frmAdviceEditEx.mint��Ч = 1
    frmAdviceEditEx.mstr�Ա� = mstr�Ա�
    frmAdviceEditEx.mlng��ĿID = FileTypeID
    frmAdviceEditEx.mstrExtData = strExtData
    frmAdviceEditEx.mint������� = PatientType + 1

    On Error Resume Next
    Me.stbThis.Panels(2).Text = "��ѡ��" + strHelpText + "..."
    frmAdviceEditEx.Show 1, Me
    Me.stbThis.Panels(2).Text = ""
    On Error GoTo errH

    If Not frmAdviceEditEx.mblnOK Then Exit Function
    If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    Me.cmd�ɼ�.Tag = rsTmp("ID")
    Me.txt�ɼ� = rsTmp("����"): Me.txt�ɼ�.Tag = Me.txt�ɼ�
    
    strAdviceText = txtҽ������.Text
    strExtData = frmAdviceEditEx.mstrExtData

    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID " + _
        "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ���� C Where A.ID=B.������ĿID And A.ID=C.������ĿID " + _
        "And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN([1],3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Len(Trim(mstr�Ա�)) = 0, "0)", IIf(mstr�Ա� Like "*��*", "1,0)", "2,0)")) + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And A.ID=[2] And B.�����ļ�ID=[3] And Ӧ�ó���=[1]"
    If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
    Set rsInput = OpenSQLRecord(strSQL, Me.Caption, PatientType + 1, Split(Split(strExtData, ";")(0), ",")(0), FileTypeID)
    
    Call AdviceSet�������(3, strExtData)
    txtҽ������.Text = Get�����������(2, "")
    strAdviceText = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
    Me.txt���� = Split(strExtData, ";")(1)
    str�걾��λ = Me.txt����
    
    '��ʼʱ��
    Me.txt��ʼʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("ִ�а���ID") = 1 Then
        Me.lbl��ʼʱ��.Visible = False: Me.chk��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Value = 0
        Me.txt��ʼʱ��.Enabled = False
    Else
        Me.lbl��ʼʱ��.Visible = True: Me.chk��ʼʱ��.Visible = False
        Me.txt��ʼʱ��.Enabled = True
    End If
    
    '����Ƶ��
    If rsInput("ִ��Ƶ��ID") = 1 Then
        Me.txtƵ��.Enabled = False: Me.txtƵ�� = "һ����": Me.cmdƵ��.Enabled = False
    Else
        Me.txtƵ��.Enabled = True: Me.txtƵ�� = "": Me.cmdƵ��.Enabled = True
    End If
    strSequence = Me.txtƵ��
    
    '����
    Me.txt���� = "1": Me.lbl������λ.Caption = rsInput("���㵥λ")
    
    '����
    If (rsInput("ִ��Ƶ��ID") = 0 And InStr(",1,2,", rsInput("���㷽ʽID")) > 0) _
                    Or InStr(",5,6,", rsInput("���ID")) > 0 Then
        Me.txt����.Enabled = True: Me.txt���� = "": Me.txt����.BackColor = Me.txtҽ������.BackColor: Me.lbl������λ.Caption = rsInput("���㵥λ")
    Else
        Me.txt����.Enabled = False: Me.txt���� = "": Me.txt����.BackColor = Me.BackColor: Me.lbl������λ.Caption = "" ' rsInput("���㵥λ")
    End If
    
    'ִ�п���
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType + 1, DeptID)
    If rsTmp Is Nothing Then
        Me.cboִ�п���.Clear: Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cboִ�п���.Clear
        Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(0) = rsTmp("ID"): Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = False: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    Else
        Me.cboִ�п���.Clear
        Do While Not rsTmp.EOF
            Me.cboִ�п���.AddItem rsTmp("����") & "-" & rsTmp("����"): Me.cboִ�п���.ItemData(Me.cboִ�п���.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cboִ�п���.ListIndex = 0
        Me.cboִ�п���.Enabled = True: Me.cboִ�п���.BackColor = Me.txtҽ������.BackColor
    End If
    
    '����ҽ��
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '����������Ŀ������ʾ��ʽ
    
    str��� = rsInput("���ID"): lngClinicID = rsInput("������ĿID"): Call ProFile1(0).SetDiagItem(lngClinicID, str�걾��λ)
    int�Ƽ����� = rsInput("�Ƽ�����ID"): intִ������ = rsInput("ִ�п���ID"): strClinicName = IIf(intType = 4, Me.txtҽ������, rsInput("����"))
    
    LabsInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '���¼��벿λ�л򸽼������м�������Ŀ��
    If int���� = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    ElseIf int���� = 3 Then
        '���������Ŀ
        strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п��� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    strDate = IIf(Me.chk��ʼʱ��.Visible And Me.chk��ʼʱ��.Value = 0, "", Format(Me.txt��ʼʱ��, "yy��MM��dd��"))
    
    If rsRelativeAdvice Is Nothing Then
        If int���� = 1 Then
            Get����������� = txtMainAdvice & IIf(Len(str�걾��λ) = 0, "", "(" & str�걾��λ & ")"): Exit Function
        Else
            Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice & IIf(Len(str�걾��λ) = 0, "", "(" & str�걾��λ & ")"): Exit Function
        End If
    End If
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If int���� = 1 Then
            If Len(Trim(rsRelativeAdvice("�걾��λ"))) > 0 Then
                strTmp = strTmp & "," & rsRelativeAdvice("�걾��λ")
            End If
        ElseIf Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") = "G" Then
                str���� = rsRelativeAdvice("����")
            Else
                strTmp = strTmp & "," & rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If int���� = 1 Then
        If strTmp <> "" Then
            Get����������� = txtMainAdvice & "(" & Mid(strTmp, 2) & ")"
        Else
            Get����������� = txtMainAdvice
        End If
    Else
        If strTmp <> "" Or str���� <> "" Then
            If str���� <> "" Then
                Get����������� = IIf(Len(strDate) = 0, "", strDate & " ") & "�� " & str���� & " ���� " & txtMainAdvice
            Else
                Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice
            End If
            If strTmp <> "" Then
                Get����������� = Get����������� & " �� " & Mid(strTmp, 2)
            End If
        Else
            Get����������� = IIf(Len(strDate) = 0, "", strDate & " �� ") & txtMainAdvice
        End If
    End If
End Function

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function

Private Function Get��������() As String
    If rsRelativeAdvice Is Nothing Then Get�������� = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            If rsRelativeAdvice("���") = "G" Then
                Get�������� = rsRelativeAdvice("����")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get��λ����() As String
    If rsRelativeAdvice Is Nothing Then Get��λ���� = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("�걾��λ"))) > 0 Then
            Get��λ���� = Get��λ���� & "," & rsRelativeAdvice("�걾��λ")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get��λ����) > 0 Then Get��λ���� = Mid(Get��λ����, 2)
End Function

Private Function GetExeDepart(ByVal lngDiagItem As Long, ByVal iPatientType As Integer, Optional ByVal lngDepartID As Long = 0) As ADODB.Recordset
'���ܣ���ȡִ�п���
'   iPatientType���������� 1=���2=סԺ
'   lngDepartID����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo DBError
    
    If lngDepartID = 0 Then lngDepartID = UserInfo.����ID
    
    strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem)
    Select Case rsTmp("ִ�п���")
        Case 0, 1, 2 '0-��ִ�еĶ�����1-�������ڿ��ң�2-�������ڲ���
            strSQL = "Select B.ID,B.����,B.���� From ������Ϣ A,���ű� B Where " & _
                IIf(rsTmp("ִ�п���") = 1, "a.��ǰ����ID", "a.��ǰ����ID") & "=B.ID And A.����ID=[1] Order by B.����"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID)
        Case 3 '���������ڿ���
            strSQL = "Select B.ID,B.����,B.���� From ���ű� B Where B.ID=[1] Order by B.����"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDepartID)
        Case 4 'ָ������
            strSQL = "Select Distinct B.ID,B.����,B.���� From ����ִ�п��� A,���ű� B Where A.������ĿID=[1]" & _
                " And A.��������ID=[2] And A.ִ�п���ID=B.ID Order by B.����"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem, lngDepartID)
            '��ѯһ�㲿��
            If rsTmp.EOF Then
                strSQL = "Select Distinct B.ID,B.����,B.���� From ����ִ�п��� A,���ű� B Where A.������ĿID=[1]" & _
                    " And ������Դ=[2] And A.ִ�п���ID=B.ID Order by B.����"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem, iPatientType)
            End If
            If rsTmp.EOF Then
                strSQL = "Select Distinct B.ID,B.����,B.���� From ����ִ�п��� A,���ű� B Where A.������ĿID=[1]" & _
                    " And A.ִ�п���ID=B.ID Order by B.����"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngDiagItem)
            End If
        Case 5 'Ժ��ִ��
            Exit Function
    End Select
    
    
    If Not rsTmp.EOF Then Set GetExeDepart = rsTmp
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupCount(lng���ID As Long) As Long
'���ܣ���ȡ�����Ŀ�е���Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(*) as NUM From ������Ŀ��� Where �������ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng���ID)
    If Not rsTmp.EOF Then GetGroupCount = zlCommFun.NVL(rsTmp!NUM, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Getȱʡ�÷�ID(int���� As Integer) As Long
'���ܣ�����ȱʡ�ĸ�ҩ;������ҩ�巨
'������int����=2-��ҩ;��,3-��ҩ�巨,4-��ҩ�÷�
'      str�Ա�=�����Ա�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From ������ĿĿ¼" & _
        " Where ���='E' And ��������=[1]" & _
        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
        " Order by ����"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", int����)
    If Not rsTmp.EOF Then Getȱʡ�÷�ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemField(ByVal lng��ĿID As Long, ByVal strField As String) As Variant
'���ܣ���ȡָ��������Ŀ��ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lng��ĿID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get����ҽ��(ByVal lng����ID As Long, ByVal bln��ʿվ As Boolean, strȱʡҽ�� As String, lngҽ��ID As Long, _
    Optional objCbo As Object, Optional ByVal int��Χ As Integer = 2) As Boolean
'���ܣ���ȡ���õĿ���ҽ����ָ������������
'������lng���˿���ID=�������ڿ���ID
'      bln��ʿվ=�Ƿ��ɻ�ʿ��ҽ����ҽ��
'      objCbo=Ҫ����ҽ���嵥��������
'      strȱʡҽ��=ȱʡ��λ��ҽ��,�������objCbo,�������ȶ�λ,�ٷ���ȱʡҽ����ҽ��ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
    
    If bln��ʿվ Then
        '�������ڿ��ҵ�ҽ��
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID=" & lng���˿���ID & _
            " Order by A.����"
        '�������ڲ������Ƶ�ҽ��
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=" & lng���˿���ID
        strSQL = "Select Distinct ����ID From ��λ״����¼ Where ����ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
        'ȫԺסԺ���ҵ�ҽ��
        strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(" & int��Χ & ",3)"
        strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & IIf(objCbo Is Nothing, ",B.����ID", "") & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            " And B.����ID IN(" & strSQL & ")" & _
            " Order by A.����"
    Else 'ҽ����ҽ��ʱ,����Ϊֻ��Ϊҽ������
        strSQL = "Select ID,���,����,���� From ��Ա�� Where ID=" & UserInfo.ID
    End If

    OpenRecord rsTmp, strSQL, "zlCISCore"
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln��ʿվ Then
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            ElseIf bln��ʿվ Then
                If strȱʡҽ�� <> "" Then
                    'ȱʡҽ��(סԺҽʦ)����
                    rsTmp.Filter = "����='" & strȱʡҽ�� & "'"
                Else
                    '���˿��ҵ�ҽ������
                    rsTmp.Filter = "����ID=" & lng���˿���ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lngҽ��ID = rsTmp!ID
                strȱʡҽ�� = rsTmp!����
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem zlCommFun.NVL(rsTmp!����) & "-" & rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!���� = strȱʡҽ�� Then
                Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    End If
    Get����ҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lng���˿���ID As Long, Optional ByVal int��Χ As Integer = 2) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1�����˿���
'      2������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      3������������/סԺ���˵Ŀ���
'      4��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 4) As Long
    
    '���ܲ���û������
    strSQL = "Select Distinct C.����,A.����ID,Nvl(A.ȱʡ,0) as ȱʡ,Nvl(B.�������,0) as �������" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID(+) And A.��ԱID=[1]" & _
        " Order by C.����"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISWork", lngҽ��ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lng���˿���ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 And rsTmp!ȱʡ = 1 Then
            arr����ID(2) = rsTmp!����ID
        ElseIf InStr("," & int��Χ & ",3,", rsTmp!�������) > 0 Then
            If arr����ID(3) = 0 Then arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(4) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowTemplate(ByVal lngElementID As Long)
'��ʾ�����ڵ�ǰԪ�ص�ģ����
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim objCurrNode As MSComctlLib.Node
    
    On Error GoTo errH
    strSQL = "Select Distinct 0 As ĩ��,�ϼ�ID,ID,����,'' As ����,���� From ����ģ�����" & _
        " Start With ID In" & _
        " (Select A.ģ�����ID From ����ģ��Ӧ�� A,����ģ����� B Where A.ģ�����ID=B.ID And ����Ԫ��ID=[1] And " & _
        "(B.������Ա Is Null Or B.������Ա='" & UserInfo.���� & "'))" & _
        " Connect By Prior �ϼ�ID=ID" & _
        " Union All" & _
        " Select 1,a.����ID,a.ID,a.����,a.����,a.���� From ����ģ������ a,����ģ��Ӧ�� b,����ģ����� c" & _
        " Where a.����id=b.ģ�����id And b.ģ�����ID=c.ID And b.����Ԫ��id=[1] And (c.������Ա Is Null Or c.������Ա='" & UserInfo.���� & "') Order By ĩ��,����"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngElementID)
    
    tvwElement.Nodes.Clear
    Do While Not rsTmp.EOF
        With tvwElement
            If IsNull(rsTmp("�ϼ�ID")) Then
                Set objCurrNode = .Nodes.Add(, , IIf(rsTmp("ĩ��") = 0, "C", "T") & rsTmp("ID"), rsTmp("����"), _
                    IIf(rsTmp("ĩ��") = 0, "Close", "Template"), IIf(rsTmp("ĩ��") = 0, "Open", "Template"))
                objCurrNode.Expanded = True
            Else
                Set objCurrNode = .Nodes.Add("C" & rsTmp("�ϼ�ID"), tvwChild, IIf(rsTmp("ĩ��") = 0, "C", "T") & rsTmp("ID"), rsTmp("����"), _
                    IIf(rsTmp("ĩ��") = 0, "Close", "Template"), IIf(rsTmp("ĩ��") = 0, "Open", "Template"))
            End If
            objCurrNode.Tag = NVL(rsTmp("����"))
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
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

