VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frm��Ӧ�̹��� 
   Caption         =   "��Ӧ�̹���"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13065
   Icon            =   "frm��Ӧ�̹���.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ils32 
      Left            =   3180
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":0D22
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":117A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":15CE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":1A26
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3300
      Top             =   3795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":1E7E
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":22D6
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":272E
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":2B82
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":2FDA
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   5985
      Left            =   2805
      TabIndex        =   4
      Top             =   720
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   10557
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���֤��"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "���֤Ч��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ִ�պ�"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ִ��Ч��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "��ַ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "�绰"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "˰��ǼǺ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "�ʺ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "��ϵ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "������"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "���ö�"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Key             =   "վ���"
         Object.Tag             =   "վ���"
         Text            =   "Ժ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Key             =   "����ʱ��"
         Object.Tag             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Key             =   "����ʱ��"
         Object.Tag             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   10557
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7815
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm��Ӧ�̹���.frx":3432
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17965
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   4800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":3CC6
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":3EE6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4106
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4322
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4542
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4762
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":497E
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4B9A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4DB4
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":4F0E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":512A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":534A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":5564
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":577E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":5998
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":5BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":5DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":5FE6
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   5520
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6200
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6420
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6640
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":685C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6A7C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6C9C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":6EB8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":70D4
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":72EE
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7448
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7668
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7888
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":7ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":80F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":830A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��Ӧ�̹���.frx":8524
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   13065
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   12810
         _ExtentX        =   22595
         _ExtentY        =   1270
         ButtonWidth     =   1455
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���ӷ���"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸ķ���"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ������"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Restore"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Description     =   "ͣ��"
               Object.ToolTipText     =   "ͣ��"
               Object.Tag             =   "ͣ��"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "StateSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
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
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "filtrate"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm��Ӧ�̹���.frx":873E
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   11400
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "����"
            Top             =   210
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   10800
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   7
            Top             =   210
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   74
               Width           =   495
            End
         End
      End
   End
   Begin MSComctlLib.ListView lvwTemp 
      Height          =   5985
      Left            =   2805
      TabIndex        =   6
      Top             =   750
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   10557
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "���֤��"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���֤Ч��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ִ�պ�"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ִ��Ч��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "��ַ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�绰"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "˰��ǼǺ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "�ʺ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "��ϵ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "������"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "���ö�"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblHsc 
      Height          =   5985
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   750
      Width           =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAddP 
         Caption         =   "���ӷ���(&P)"
      End
      Begin VB.Menu mnuEditUpdateP 
         Caption         =   "�޸ķ���(&U)"
      End
      Begin VB.Menu mnuEditDeleteP 
         Caption         =   "ɾ������(&D)"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "������Ŀ(&A)"
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "�޸���Ŀ(&X)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ����Ŀ(&B)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&S)"
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "����(&R)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
         End
         Begin VB.Menu mnuViewLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHide 
         Caption         =   "��ʾͣ����Ŀ(&H)"
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFiltrate 
         Caption         =   "����(&I)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewASP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&T)"
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
            Caption         =   "���ͷ���(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)"
      End
   End
   Begin VB.Menu mnuFast 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuFastAdd 
         Caption         =   "������Ŀ(&A)"
      End
      Begin VB.Menu mnuFastModify 
         Caption         =   "�޸���Ŀ(&E)"
      End
      Begin VB.Menu mnuFastDelete 
         Caption         =   "ɾ����Ŀ(&D)"
      End
      Begin VB.Menu mnuFastLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFastRestore 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuFastStop 
         Caption         =   "ͣ��(&S)"
      End
      Begin VB.Menu mnuFastLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "��ϸ����(&T)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm��Ӧ�̹���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownX As Single, msngDownY As Single, mSaveKey As String, mFocus As Integer, mstrFilt As String, mintFilt As Integer
Private mrstFind As New ADODB.Recordset, mFirstID As String, mLastID As String, mintColumn As Integer
Private mcllFilter As Collection
Private mblnFirst As Boolean
Private mstrPrivs As String
Dim mstrĬ��Ȩ�� As String
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String

Private Sub Form_Activate()
    Call Form_Resize
    If mblnFirst = False Then Exit Sub
    mSaveKey = ""
    mblnFirst = False
    
    Call InitFilter
    'Ȩ������
    Call Ȩ�޿���
     
    '����������
    Call FullType
    '������ϸ����
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Integer
    mstrPrivs = gstrPrivs
    mblnFirst = True
    
    RestoreWinState Me, App.ProductName
    
    mnuViewIcon(lvwList.View).Checked = True
    mnuFastIcon(lvwList.View).Checked = True
    lvwList.Sorted = False
    
    Call InitFilter
    mstrĬ��Ȩ�� = GetDefault����
    
    Err = 0
    On Error Resume Next
    mstrFilt = ""
    For i = 1 To Len(mstrĬ��Ȩ��)
        If Mid(mstrĬ��Ȩ��, i, 1) = 1 Then
            mstrFilt = mstrFilt & " or substr(����," & i & ",1)=1"
        End If
    Next
    If mstrFilt <> "" Then
        mstrFilt = "  ( " & Mid(mstrFilt, 4) & " )"
    End If
   
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub mnuEditDel_Click()
    Call mnuEditDelete_Click
End Sub

Private Sub mnuEditDeleteP_Click()
    Call mnuEditDelete_Click
End Sub

Private Sub mnuEditUpdate_Click()
    Call mnuEditModify_Click
End Sub

Private Sub mnuEditUpdateP_Click()
    Call mnuEditModify_Click
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����id As Long
    Dim lng��Ӧ��ID As Long
    'Dim byt����ͣ�� As Byte
    
    'byt����ͣ�� = IIf(mnuViewHide.Checked, 1, 0)
    If Not tvwList.SelectedItem Is Nothing Then
        lng����id = Val(Mid(Me.tvwList.SelectedItem.Key, 2))
    End If
    
    If Not lvwList.SelectedItem Is Nothing Then
        lng��Ӧ��ID = Val(Mid(lvwList.SelectedItem.Key, 2))
    End If
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "����=" & lng����id, "��Ӧ��=" & lng��Ӧ��ID)
    
End Sub

Private Sub InitFilter()
    Set mcllFilter = New Collection
    mcllFilter.Add Array("", ""), "����"
    mcllFilter.Add "", "����"
    mcllFilter.Add Array("0", "0"), "������"
    mcllFilter.Add Array("0", "0"), "���ö�"
End Sub

Private Sub FullType()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:װ�빩Ӧ�̷���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim TmpNode As Node
    
    gstrSQL = "" & _
        "   Select ID,�ϼ�ID,����,���� " & _
        "   From ��Ӧ��  " & _
        "   Where ĩ�� <> 1 " & _
        "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    
    Err = 0
    
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    tvwList.Nodes.Clear
    Set TmpNode = tvwList.Nodes.Add(, , "Root", "���й�Ӧ��", 1, 1)
    
    TmpNode.Sorted = True
    TmpNode.Expanded = True
    TmpNode.Selected = True
    
    Do While Not rsTemp.EOF
        If IsNull(rsTemp!�ϼ�ID) Then
            Set TmpNode = tvwList.Nodes.Add("Root", 4, "K" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����, 5, 5)
        Else
            Set TmpNode = tvwList.Nodes.Add("K" & rsTemp!�ϼ�ID, 4, "K" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����, 5, 5)
        End If
        TmpNode.Sorted = True
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    If cbrThis.Bands(1).MinHeight <> tlbThis.Height Then cbrThis.Bands(1).MinHeight = tlbThis.Height
    
    cbrThis.Move 0, 0, Me.ScaleWidth
    
    If lblHsc.Left > Me.ScaleWidth - 2000 Then lblHsc.Left = Me.ScaleWidth - 2000
    
    lblHsc.Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
    lblHsc.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc.Top
    
    tvwList.Move 0, lblHsc.Top, lblHsc.Left, lblHsc.Height
    lvwList.Move lblHsc.Left + lblHsc.Width, lblHsc.Top, Me.ScaleWidth - (lblHsc.Left + lblHsc.Width), lblHsc.Height
    lvwTemp.Move lvwList.Left, lvwList.Top, lvwList.Width, lvwList.Height

    mnuViewToolButton.Checked = cbrThis.Visible
    mnuViewStatus.Checked = stbThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lvwList.Sorted = False
    mstrFindValue = ""
    Set mrsFind = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblHsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
    End If
End Sub

Private Sub lvwList_Click()
    If Me.lvwList.SelectedItem Is Nothing Then
        SetEnabled
    End If
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwList.Sorted = True
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwList.SortOrder = IIf(lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwList.SortKey = mintColumn
        lvwList.SortOrder = lvwAscending
    End If
'    lvwList.Refresh
End Sub

Private Sub lvwList_DblClick()
    Dim blnȨ�� As Boolean
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    blnȨ�� = SetEditPro(Split(lvwList.SelectedItem.Tag, "|")(1))
    If Me.mnuEdit.Visible = False Or Me.mnuEditModify.Visible = False Or blnȨ�� = False Then
        '�ɲ鿴
        frm��Ӧ�̱༭.�༭��λ Me, Val(lvwList.SelectedItem.Tag), g�鿴, Mid(lvwList.SelectedItem.Key, 2), True
    Else
        mnuEditModify_Click
    End If
End Sub

Private Sub lvwList_GotFocus()
    mFocus = 2
    SetEnabled
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SetEnabled
End Sub

Private Sub lvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
    If Button = 2 Then
        PopupMenu mnuFast
    End If
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Ȩ�޿���
'    If Button = 2 Then
'        mnuFastLine1.Visible = True
'        mnuFastLine2.Visible = True
'        mnuFastStop.Visible = True
'        mnuFastRestore.Visible = True
'        mnuFastIcon(0).Visible = True
'        mnuFastIcon(1).Visible = True
'        mnuFastIcon(2).Visible = True
'        mnuFastIcon(3).Visible = True
'        Me.PopupMenu mnuFast
'    End If
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnReturn As Boolean
    Dim strLstKey As String
    
    blnReturn = frm��Ӧ�̱༭.�༭��λ(Me, Val(Mid(tvwList.SelectedItem.Key, 2)), g����, "", True, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Err = 0
    On Error Resume Next
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    '�ָ�ѡ��
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    
    mSaveKey = ""
    tvwList_NodeClick tvwList.SelectedItem
    '�ָ���ʷѡ������
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddP_Click()
    Dim blnReturn As Boolean
    Dim strSaveKey As String
    Dim strLstKey As String
    blnReturn = frm��Ӧ�̱༭.�༭��λ(Me, Val(Mid(tvwList.SelectedItem.Key, 2)), g����, "", False, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Err = 0
    On Error Resume Next
    strSaveKey = tvwList.SelectedItem.Key
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    mSaveKey = ""
    '����װ������
    Call FullType
    '�ָ�ѡ��
    tvwList.Nodes(strSaveKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList.SelectedItem.Expanded = True
    tvwList_NodeClick tvwList.SelectedItem
    '�ָ���ʷѡ������
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Long
    Dim strSQL As String
    Dim blnYes As Boolean
    Dim blnActTree As Boolean
    Dim mstrKey As String
    blnActTree = Me.ActiveControl Is tvwList
    
    If blnActTree Then
        If Me.tvwList.SelectedItem Is Nothing Then Exit Sub
        ShowMsgbox "��ȷ��Ҫɾ������(�����¼���Ŀ)Ϊ" & vbCrLf & "��" & Me.tvwList.SelectedItem.Text & "���ļ�¼��", True, blnYes
        mstrKey = Me.tvwList.SelectedItem.Key
    Else
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        If SetEditPro(Split(Me.lvwList.SelectedItem.Tag, "|")(1)) = False Then Exit Sub
        
        ShowMsgbox "��ȷ��Ҫɾ����Ӧ��Ϊ" & vbCrLf & "��" & Me.tvwList.SelectedItem.Text & "���ļ�¼��", True, blnYes
        mstrKey = Me.lvwList.SelectedItem.Key
    End If
    If blnYes = False Then Exit Sub
    
    If ActiveControl Is tvwList Then
        strSQL = "zl_��Ӧ��_delete(" & Mid(tvwList.SelectedItem.Key, 2) & ")"
    Else
        strSQL = "zl_��Ӧ��_delete(" & Mid(lvwList.SelectedItem.Key, 2) & ")"
    End If
    Err = 0
    On Error GoTo errHandle:
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
    If blnActTree Then
        If tvwList.SelectedItem.Next Is Nothing Then
            If tvwList.SelectedItem.Previous Is Nothing Then
                tvwList.Nodes.Remove mstrKey
            Else
                Set tvwList.SelectedItem = tvwList.SelectedItem.Previous
                tvwList.Nodes.Remove mstrKey
            End If
        Else
            tvwList.SelectedItem.Next.Selected = True
            tvwList.Nodes.Remove mstrKey
        End If
        mSaveKey = tvwList.SelectedItem.Key
        FullList
    Else
        With lvwList
            '��ɾ��ListView�ж�Ӧ�ڵ�
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            Else
                .SetFocus
            End If
        End With
    End If
    
    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    Dim blnReturn  As Boolean
    Dim strLstKey As String
    Dim blnĩ�� As Boolean
    Dim strID As String
    Dim lng�ϼ�id As Long
    
    blnĩ�� = Me.ActiveControl Is lvwList
    If blnĩ�� Then
        If lvwList.SelectedItem Is Nothing Then Exit Sub
        strID = Mid(Me.lvwList.SelectedItem.Key, 2)
        lng�ϼ�id = Val(Split(Me.lvwList.SelectedItem.Tag, "|")(0))
        
        If SetEditPro(Split(Me.lvwList.SelectedItem.Tag, "|")(1)) = False Then Exit Sub
        
        blnReturn = frm��Ӧ�̱༭.�༭��λ(Me, lng�ϼ�id, g�޸�, strID, blnĩ��, mstrPrivs)
        
        If blnReturn = False Then Exit Sub
        
        Err = 0
        On Error Resume Next
        If lvwList.SelectedItem Is Nothing Then
            strLstKey = ""
        Else
            strLstKey = lvwList.SelectedItem.Key
        End If
        '�ָ�ѡ��
        If tvwList.SelectedItem Is Nothing Then
            tvwList.Nodes("Root").Selected = True
            tvwList.Nodes("Root").Expanded = True
        End If
        mSaveKey = ""
        tvwList_NodeClick tvwList.SelectedItem
        '�ָ���ʷѡ������
        lvwList.ListItems(strLstKey).Selected = True
        lvwList.ListItems(strLstKey).EnsureVisible
        Err = 0
        On Error GoTo 0
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    If tvwList.SelectedItem.Key = "Root" Then Exit Sub
    If tvwList.SelectedItem Is Nothing Then Exit Sub
    strID = Mid(tvwList.SelectedItem.Key, 2)
    blnReturn = frm��Ӧ�̱༭.�༭��λ(Me, Val(Mid(tvwList.SelectedItem.Parent.Key, 2)), g�޸�, strID, blnĩ��, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Dim strSaveKey  As String
    
    Err = 0
    On Error Resume Next
    strSaveKey = tvwList.SelectedItem.Key
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    mSaveKey = ""
    '����װ������
    Call FullType
    '�ָ�ѡ��
    tvwList.Nodes(strSaveKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
    '�ָ���ʷѡ������
    
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditRestore_Click()
    Dim strSQL As String
    
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    strSQL = "zl_��Ӧ��_reuse (" & Mid(lvwList.SelectedItem.Key, 2) & ")"
        
    On Error GoTo errHandle:
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    lvwList.SelectedItem.Icon = 2
    lvwList.SelectedItem.SmallIcon = 2
    lvwList.SelectedItem.SubItems(18) = ""

    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim strSQL As String
    Dim intIndex As Integer
    
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        
    strSQL = "zl_��Ӧ��_stop(" & Mid(lvwList.SelectedItem.Key, 2) & ")"
        
    On Error GoTo errHandle:
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If mnuViewHide.Checked Then
        lvwList.SelectedItem.Icon = 3
        lvwList.SelectedItem.SmallIcon = 3
        lvwList.SelectedItem.SubItems(18) = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        
    Else
        With lvwList
            '��ɾ��ListView�ж�Ӧ�ڵ�
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            Else
                .SetFocus
            End If
        End With
    End If
    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuFastAdd_Click()
    mnuEditAdd_Click
End Sub

Private Sub mnuFastChild_Click()
    mnuEditAddP_Click
End Sub

Private Sub mnuFastDelete_Click()
    mnuEditDelete_Click
End Sub

Private Sub mnuFastIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuFastModify_Click()
    mnuEditModify_Click
End Sub

Private Sub mnuFastRestore_Click()
    mnuEditRestore_Click
End Sub

Private Sub mnuFastStop_Click()
    mnuEditStop_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    subPrint 2
End Sub

Private Sub mnuViewFiltrate_Click()
    Dim blnCancel As Boolean
    Dim strFilter As String
    Dim cllFilter As Collection
    
    Dim intOldCondition As Integer
    intOldCondition = mintFilt
    Call frm��Ӧ�̹���.GetFiler(Me, blnCancel, mstrFilt, cllFilter, mstrPrivs)
    If blnCancel = True Then Exit Sub
    If mstrFilt <> "" Then
        mstrFilt = "(" & mstrFilt & ")"
    End If
    Set mcllFilter = cllFilter
    FullList
End Sub

Private Sub mnuViewFind_Click()
    '���ҹ���
    Dim strSQL As String
    Dim strOthers() As String

    strSQL = frm��Ӧ�̶�λ.getSql(strOthers)
    If strSQL = "" Then Exit Sub
    strSQL = strSQL & IIf(mnuViewHide.Checked, "", " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))") & IIf(mstrFilt <> "", " And " & mstrFilt, "")
    Set mrstFind = New ADODB.Recordset
'    zlDatabase.OpenRecordset mrstFind, strSql, Me.Caption
    On Error GoTo errHandle
    Set mrstFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(mcllFilter("����")(0)), CStr(mcllFilter("����")(1)), _
                            CStr(mcllFilter("����")), CLng(mcllFilter("������")(0)), CLng(mcllFilter("������")(1)), _
                            CDbl(mcllFilter("���ö�")(0)), CDbl(mcllFilter("���ö�")(1)), strOthers(0), strOthers(1), strOthers(2))


    mrstFind.Sort = "�ϼ�ID,����,ID"
    If mrstFind.EOF Then
        MsgBox "û�����㶨λ���������ݣ�", vbInformation, Me.Caption
        Exit Sub
    End If
    mrstFind.MoveFirst
    mFirstID = mrstFind("ID")
    mrstFind.MoveLast
    mLastID = mrstFind("ID")
    Unload frm��Ӧ�̶�λ
    frmToolBarWin.ShowBar "��Ӧ�̶�λ", Me
    subFirst
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Find��Ӧ��()
    Dim strKeytvw As String, strKeylvw As String, blnUP As Boolean, blnDown As Boolean
    Dim rstTemp As New ADODB.Recordset
    If mrstFind.EOF Then Exit Sub
    frmToolBarWin.���� 0, False
    frmToolBarWin.���� 1, False
    On Error GoTo errHandle
    If Not IsNull(mrstFind("�ϼ�ID")) Then
        'by lesfeng 2009-12-2 �����Ż�
        Set rstTemp = zlDatabase.OpenSQLRecord("Select id From ��Ӧ�� Where ID=[1]", Me.Caption, Val(mrstFind!�ϼ�ID))
        If Not rstTemp.EOF Then
            strKeytvw = "K" & rstTemp!ID
        End If
        rstTemp.Close
    Else
        strKeytvw = "Root"
    End If
    strKeylvw = "K" & mrstFind("ID")
    If strKeytvw = mSaveKey Then
        Set lvwList.SelectedItem = lvwList.ListItems(strKeylvw)
    Else
        Set tvwList.SelectedItem = tvwList.Nodes(strKeytvw)
        FullList
        Set lvwList.SelectedItem = lvwList.ListItems(strKeylvw)
    End If
    blnUP = (mrstFind("ID") <> mFirstID)
    blnDown = (mrstFind("ID") <> mLastID)
    frmToolBarWin.���� 0, blnUP
    frmToolBarWin.���� 1, blnDown
    tvwList.SelectedItem.EnsureVisible
    lvwList.SelectedItem.EnsureVisible
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub subFirst()
    mrstFind.MoveFirst
    Find��Ӧ��
End Sub

Public Sub subPrevious()
    mrstFind.MovePrevious
    Find��Ӧ��
End Sub

Public Sub subNext()
    mrstFind.MoveNext
    Find��Ӧ��
End Sub

Public Sub subLast()
    mrstFind.MoveLast
    Find��Ӧ��
End Sub

Private Sub mnuViewHide_Click()
    mnuViewHide.Checked = Not mnuViewHide.Checked
    FullList
    Set mrsFind = Nothing
    mstrFindValue = ""
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
        mnuFastIcon(intTemp).Checked = False
    Next
    
    mnuViewIcon(Index).Checked = True
    mnuFastIcon(Index).Checked = True
    lvwList.View = Index
    lvwList.Refresh
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    mSaveKey = ""
    If Me.tvwList.SelectedItem Is Nothing Then
        strKey = "Root"
    Else
        strKey = Me.tvwList.SelectedItem.Key
    End If
    Call FullType
    Err = 0
    On Error Resume Next
    tvwList.Nodes(strKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    Else
        tvwList.SelectedItem.Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub



Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lvwTemp As ListView
    Select Case Button.Key
        Case "filtrate"
            mnuViewFiltrate_Click
        Case "Add"
            If Me.ActiveControl Is tvwList Then
                mnuEditAddP_Click
            Else
                mnuEditAdd_Click
            End If
        Case "Modify"
            mnuEditModify_Click
        Case "Print"
            mnuFilePrint_Click
        Case "PrintView"
            mnuFilePrintView_Click
        Case "Find"
            mnuViewFind_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Restore"
            mnuEditRestore_Click
        Case "Stop"
            mnuEditStop_Click
        Case "View"
            Set lvwTemp = lvwList
            mnuViewIcon(lvwTemp.View).Checked = False
            mnuFastIcon(lvwTemp.View).Checked = False
            If lvwTemp.View = 3 Then
                mnuViewIcon(0).Checked = True
                mnuFastIcon(0).Checked = True
                lvwTemp.View = 0
            Else
                mnuViewIcon(lvwTemp.View + 1).Checked = True
                mnuFastIcon(lvwTemp.View + 1).Checked = True
                lvwTemp.View = lvwTemp.View + 1
            End If
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call mnuViewIcon_Click(ButtonMenu.Index - 1)
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Parent Is Nothing Then
        Node.Expanded = True
        Exit Sub
    End If
    If InStr(tvwList.SelectedItem.Key, Node.Key) > 0 Then
        Set tvwList.SelectedItem = Node
        tvwList_NodeClick Node
    End If
End Sub

Private Sub tvwList_GotFocus()
    mFocus = 1
    SetEnabled
End Sub

Private Sub tvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Ȩ�޿���
    If mnuEdit.Visible = False Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub tvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DoEvents
'    If Button = 2 Then
'        mnuFastLine1.Visible = False
'        mnuFastLine2.Visible = False
'        mnuFastStop.Visible = False
'        mnuFastRestore.Visible = False
'        mnuFastIcon(0).Visible = False
'        mnuFastIcon(1).Visible = False
'        mnuFastIcon(2).Visible = False
'        mnuFastIcon(3).Visible = False
'        Me.PopupMenu mnuFast
'    End If
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mSaveKey Then Exit Sub
    Call FullList
End Sub

Public Sub FullList(Optional strCon As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ϸ����
    '--�����:strCon -����
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lstItem As ListItem, strTempKey As String
    Dim strWhere As String
    
    Dim strTvwKey As String
    
    strTvwKey = tvwList.SelectedItem.Key
    
    strWhere = ""
    If strCon <> "" Then
        strWhere = " and (" & strCon & ") "
    End If
    'by lesfeng 2009-12-2 �����Ż�
    If mnuViewHide.Checked = False Then
        If strTvwKey = "Root" Then
            strWhere = strWhere & "  and (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null)"
        Else
            strWhere = strWhere & "  and (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null)" & " start with  �ϼ�ID = [8] connect by prior id=�ϼ�id  "
        End If
    Else
        If strTvwKey <> "Root" Then
            strWhere = strWhere & " start with  �ϼ�ID = [8] connect by prior id=�ϼ�id  "
        End If
    End If
    Err = 0
    On Error GoTo ErrHand:
    'by lesfeng 2009-12-2 �����Ż�
    gstrSQL = "" & _
        "   Select ID,�ϼ�ID,����,����,����,ĩ��,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,�绰,��������," & _
        "          �ʺ�,��ϵ��,����ʱ��,����ʱ��,����,������,���ö�,����ί����,����ί������,������֤��,������֤����," & _
        "          ҩ��ֱ�����,ҩ��ֱ�������,��Ȩ��,��Ȩ��,վ��" & _
        "    from ��Ӧ��  where ĩ��=1  " & IIf(mstrFilt = "", "", " And " & mstrFilt) & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mcllFilter("����")(0)), CStr(mcllFilter("����")(1)), _
                            CStr(mcllFilter("����")), CLng(mcllFilter("������")(0)), CLng(mcllFilter("������")(1)), _
                            CDbl(mcllFilter("���ö�")(0)), CDbl(mcllFilter("���ö�")(1)), Val(Mid(strTvwKey, 2)))
    
    Dim strTmp As String
    Dim i As Integer
    Dim str���� As String
    
    If lvwList.SelectedItem Is Nothing Then
        strTempKey = ""
    Else
        strTempKey = lvwList.SelectedItem.Key
    End If
    lvwList.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If Format(!����ʱ��, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!����ʱ��) Then
                Set lstItem = lvwList.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("����"), 2, 2)
            Else
                Set lstItem = lvwList.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("����"), 3, 3)
            End If
            lstItem.Tag = Nvl(!�ϼ�ID, 0) & "|" & Nvl(!����)
            
            strTmp = Nvl(!����) 'Right(Dec2Bin(Nvl(!����, 0)), 4)
            str���� = ""
            For i = 1 To Len(strTmp)
                If Mid(strTmp, i, 1) = 1 Then
                    str���� = str���� & "," & Switch(i = 1, "ҩƷ", i = 2, "����", i = 3, "�豸", i = 4, "����", i = 5, "��������")
                End If
            Next
            If str���� <> "" Then
                str���� = Mid(str����, 2)
            End If
            
            lstItem.ListSubItems.Add , , Nvl(!����)
            lstItem.ListSubItems.Add , , Nvl(!����)
            lstItem.ListSubItems.Add , , str����
            
            lstItem.ListSubItems.Add , , Nvl(!���֤��)
            lstItem.ListSubItems.Add , , Nvl(!���֤Ч��)
            lstItem.ListSubItems.Add , , Nvl(!ִ�պ�)
            lstItem.ListSubItems.Add , , Nvl(!ִ��Ч��)
            lstItem.ListSubItems.Add , , Nvl(!��ַ)
            lstItem.ListSubItems.Add , , Nvl(!�绰)
            lstItem.ListSubItems.Add , , Nvl(!˰��ǼǺ�)
            lstItem.ListSubItems.Add , , Nvl(!��������)
            lstItem.ListSubItems.Add , , Nvl(!�ʺ�)
            lstItem.ListSubItems.Add , , Nvl(!��ϵ��)
            lstItem.ListSubItems.Add , , Nvl(!������, " ")
            lstItem.ListSubItems.Add , , Nvl(!���ö�, " ")
            lstItem.ListSubItems.Add , , Nvl(!վ��, " ")
            lstItem.ListSubItems.Add , , Format(!����ʱ��, "yyyy-mm-dd")
            If Format(!����ʱ��, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!����ʱ��) Then
                lstItem.ListSubItems.Add , , " "
            Else
                lstItem.ListSubItems.Add , , Format(!����ʱ��, "yyyy-mm-dd")
            End If
            rsTemp.MoveNext
        Loop
    End With
    mSaveKey = tvwList.SelectedItem.Key
    
    If lvwList.ListItems.Count > 0 Then
        If strTempKey <> "" Then
            On Error Resume Next
            Set lvwList.SelectedItem = lvwList.ListItems(strTempKey)
        End If
        If lvwList.SelectedItem Is Nothing Then
            Set lvwList.SelectedItem = lvwList.ListItems(1)
        End If
        Err = 0
        On Error GoTo 0
    End If
    SetEnabled
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetEnabled()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����׳̬
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim blnData As Boolean '��������
    Dim blnActTree As Boolean   '��ǰ����ؼ�Ϊ��
    Dim blnRoot As Boolean      '�Ƿ�ѡ��ĸ�
    Dim blnItmSel As Boolean    '�Ƿ�ѡ��
    Dim blnStop As Boolean      'ͣ�ò���
    Dim blnChild As Boolean     '��ֻ��Ŀ¼
    Dim str���� As String
    Dim blnȨ�� As Boolean
    blnActTree = Me.ActiveControl Is tvwList
    blnData = Me.lvwList.ListItems.Count <> 0
    blnRoot = Me.tvwList.SelectedItem.Key = "Root"
    blnȨ�� = False
    If Not Me.lvwList.SelectedItem Is Nothing Then
        str���� = Split(lvwList.SelectedItem.Tag, "|")(1)
        blnȨ�� = SetEditPro(str����)
        blnStop = Trim(lvwList.SelectedItem.SubItems(18)) <> "" And blnȨ��
        blnItmSel = blnȨ��
    Else
        blnItmSel = False
    End If
    
    If blnRoot Then
        mnuEditAddP.Enabled = (blnRoot) And (blnActTree)
    Else
        mnuEditAddP.Enabled = (blnActTree)
    End If
    mnuEditUpdateP.Enabled = (Not blnRoot) And (blnActTree)
    mnuEditDeleteP.Enabled = (Not blnRoot) And (blnActTree)
    mnuEditAdd.Enabled = Not blnActTree
    mnuEditModify.Enabled = Not blnActTree
    mnuEditUpdate.Enabled = Not blnActTree
    mnuEditDelete.Enabled = Not blnActTree
    mnuEditDel.Enabled = Not blnActTree
    mnuEditStop.Enabled = (Not blnActTree) And blnItmSel And (Not blnStop)
    mnuEditRestore.Enabled = (Not blnActTree) And blnItmSel And blnStop
    
    mnuFastAdd.Enabled = mnuEditAdd.Enabled
    mnuFastDelete.Enabled = mnuEditDelete.Enabled
    mnuFastModify.Enabled = mnuEditModify.Enabled
    mnuFastStop.Enabled = mnuEditStop.Enabled
    mnuFastRestore.Enabled = mnuEditRestore.Enabled
    
    If blnActTree Then
        tlbThis.Buttons("Add").Enabled = mnuEditAddP.Enabled: tlbThis.Buttons("Add").Caption = "���ӷ���"
        tlbThis.Buttons("Modify").Enabled = mnuEditUpdateP.Enabled: tlbThis.Buttons("Modify").Caption = "�޸ķ���"
        tlbThis.Buttons("Delete").Enabled = mnuEditDeleteP.Enabled: tlbThis.Buttons("Delete").Caption = "ɾ������"
    Else
        tlbThis.Buttons("Add").Enabled = mnuEditAdd.Enabled: tlbThis.Buttons("Add").Caption = "������Ŀ"
        tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled: tlbThis.Buttons("Modify").Caption = "�޸���Ŀ"
        tlbThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled: tlbThis.Buttons("Delete").Caption = "ɾ����Ŀ"
    End If
'    tlbThis.Buttons("Add").Visible = False
'    tlbThis.Buttons("Modify").Visible = False
'    tlbThis.Buttons("Delete").Visible = False
    tlbThis.Buttons("PrintSeparate").Visible = True
    
    tlbThis.Buttons("Restore").Enabled = mnuEditRestore.Enabled
    tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    
    
    mnuFilePrint.Enabled = blnData
    mnuFilePrintView.Enabled = blnData
    mnuFileExcel.Enabled = blnData
        
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "��Ӧ���б�"
    Set objPrint.Body.objData = lvwList
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")

    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub Ȩ�޿���()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnStart As Boolean
    Dim blnStop As Boolean
    blnAdd = InStr(mstrPrivs, ";����;") <> 0
    blnModify = InStr(mstrPrivs, ";�޸�;") <> 0
    blnDelete = InStr(mstrPrivs, ";ɾ��;") <> 0
    blnStart = InStr(mstrPrivs, ";����;") <> 0
    blnStop = InStr(mstrPrivs, ";ͣ��;") <> 0
    
    If blnAdd = False And blnModify = False And blnStart = False And blnDelete = False And blnStop = False Then
        mnuEdit.Visible = False
        '��ݲ˵�
        mnuFastAdd.Visible = False
        mnuFastModify.Visible = False
        mnuFastDelete.Visible = False
        mnuFastLine1.Visible = False
        mnuFastRestore.Visible = False
        mnuFastStop.Visible = False
        mnuFastLine2.Visible = False
    Else
        mnuEditAdd.Visible = blnAdd: mnuEditAddP.Visible = blnAdd   '����
        mnuEditUpdate.Visible = blnModify: mnuEditUpdateP.Visible = blnModify   '�޸�
        mnuEditDel.Visible = blnDelete: mnuEditDeleteP.Visible = blnDelete   'ɾ��
        mnuEditLine1.Visible = (blnAdd Or blnModify Or blnDelete) And (blnStop Or blnStart)
        mnuEditRestore.Visible = blnStart
        mnuEditStop.Visible = blnStop
        '��ݲ˵�
        mnuFastAdd.Visible = blnAdd
        mnuFastModify.Visible = blnModify
        mnuFastDelete.Visible = blnDelete
        If blnAdd = False And blnModify = False And blnDelete = False Then
            mnuFastLine1.Visible = False
        End If
        mnuFastRestore.Visible = blnStart
        mnuFastStop.Visible = blnStop
        If blnStart = False And blnStop = False Then
            mnuFastLine2.Visible = False
        End If
    End If
    tlbThis.Buttons("Add").Visible = blnAdd
    tlbThis.Buttons("Modify").Visible = blnModify
    tlbThis.Buttons("Delete").Visible = blnDelete
    tlbThis.Buttons("EditSeparate").Visible = blnAdd Or blnModify Or blnDelete
    tlbThis.Buttons("StateSeparate").Visible = blnStart Or blnStop
    tlbThis.Buttons("Restore").Visible = blnStart
    tlbThis.Buttons("Stop").Visible = blnStop
    
    mnuEditModify.Visible = False
    mnuEditDelete.Visible = False
End Sub
Private Function SetEditPro(ByVal str���� As String) As Boolean
    '���ñ༭Ȩ��
    
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    
    blnҩƷ = InStr(1, mstrPrivs, "ҩƷ��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���ʹ�Ӧ��") <> 0
    bln�豸 = InStr(1, mstrPrivs, "�豸��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "������Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���Ĺ�Ӧ��") <> 0
    
    Err = 0: On Error GoTo ErrHand:
    
    SetEditPro = False
    If blnҩƷ = False And bln���� = False And bln�豸 = False And bln���� = False And bln���� = False Then
            Exit Function
    End If
    If Mid(str����, 1, 1) = 1 Then
        If Not blnҩƷ Then
            Exit Function
        End If
    End If
    
    If Mid(str����, 2, 1) = 1 Then
        If Not bln���� Then
            Exit Function
        End If
    End If
    If Mid(str����, 3, 1) = 1 Then
        If Not bln�豸 Then
            Exit Function
        End If
    End If
    
    If Mid(str����, 4, 1) = 1 Then
        If Not bln���� Then
            Exit Function
        End If
    End If
    If Mid(str����, 5, 1) = 1 Then
        If Not bln���� Then
            Exit Function
        End If
    End If
    
    
    SetEditPro = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function GetDefault����() As String
    '���ñ༭Ȩ��
    
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    Dim strTemp As String
    
    blnҩƷ = InStr(1, mstrPrivs, "ҩƷ��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���ʹ�Ӧ��") <> 0
    bln�豸 = InStr(1, mstrPrivs, "�豸��Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "������Ӧ��") <> 0
    bln���� = InStr(1, mstrPrivs, "���Ĺ�Ӧ��") <> 0
    
    strTemp = ""
    strTemp = strTemp & IIf(blnҩƷ, "1", "0")
    strTemp = strTemp & IIf(bln����, "1", "0")
    strTemp = strTemp & IIf(bln�豸, "1", "0")
    strTemp = strTemp & IIf(bln����, "1", "0")
    strTemp = strTemp & IIf(bln����, "1", "0")
    GetDefault���� = strTemp
    
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    zlCommFun.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        On Error GoTo errHandle
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null ) "
            gstrSQL = "select id,�ϼ�id from ��Ӧ�� where ���� like [1] or ���� like [1] or ���� like [1] and ĩ��=1"
            
            If mnuViewHide.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "��Ӧ�̲�ѯ", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " û���ҵ�������������Ϣ��", vbInformation, gstrSysName
        txtFind.SetFocus
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " �Ѿ���λ�������ҵ�����Ϣ������������������", vbInformation, gstrSysName
        txtFind.SetFocus
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
        Exit Sub
    End If
    
    With tvwList
        If IsNull(mrsFind("�ϼ�ID")) = False Then
            .Nodes("K" & mrsFind("�ϼ�ID")).Selected = True
        Else
            .Nodes("Root").Selected = True
        End If
        .SelectedItem.EnsureVisible
    End With
        
    With lvwList
        .ListItems("K" & mrsFind("id")).Selected = True
        .SelectedItem.EnsureVisible
    End With
End Sub
