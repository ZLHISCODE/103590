VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyList 
   Caption         =   "һ�շ����嵥"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   -495
   ClientWidth     =   8880
   Icon            =   "frmDailyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   6675
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   855
      Width           =   1410
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3855
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "���ڣ�2000��10��20��"
      Top             =   885
      Width           =   2475
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4785
      Left            =   3870
      TabIndex        =   4
      Top             =   1170
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   8440
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   5160
      Left            =   15
      TabIndex        =   0
      Top             =   795
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   9102
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483628
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "��Ա��"
         Text            =   "סԺ��"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�ѱ�"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "�Ա�"
         Text            =   "�Ա�"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "��Ժ����"
         Text            =   "��Ժ����"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "��Ժ����"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "��������"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta״̬ 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDailyList.frx":030A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8281
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   8880
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   3705
      NewRow1         =   0   'False
      Caption2        =   "���˲���"
      Child2          =   "cbo����"
      MinHeight2      =   300
      Width2          =   1995
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   4110
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   2145
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":11F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1626
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1840
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":258E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   1410
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":27A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":29C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":2BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":2DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3016
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3230
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":344A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3882
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":4198
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic�ָ� 
      Height          =   5340
      Left            =   3510
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5340
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3255
      Top             =   5385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":43B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":46CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3210
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":4FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":52C0
            Key             =   ""
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
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
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��������(&J)"
      End
      Begin VB.Menu mnuFileLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewQuitFee 
         Caption         =   "��ʾ�˷�(&Q)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewZero 
         Caption         =   "��ʾ�����(&Z)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuViewFindLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "������һ��(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "С����"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewALLSele 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuViewALLClear 
         Caption         =   "ȫ��(&C)"
         Shortcut        =   ^C
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
      Begin VB.Menu mnuHelp_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_about 
         Caption         =   "����(&A)��"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopDisp 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDailyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlngModul As Long

Private mrsPati As New ADODB.Recordset
Private mintBedLen As Integer
Private mblnPrint As Boolean '���嵥�Ƿ��ӡ
Private mdtMin As Date, mdtMax As Date
Private mbln��ҽ������ As Boolean
Private mblnҽ������ As Boolean
Private mbln��Ժ���� As Boolean
Private mbln��Ժ���� As Boolean
Private mstr����ʱ�� As String
Private mbyt���˲���ģʽ As Byte '0-�з��õĲ���(ȱʡ)��1-���˵�ǰ����

Private Sub cboPage_Change()
    Refresh�����嵥 lvwPati.SelectedItem, Val(cboPage.ItemData(cboPage.ListIndex))
End Sub

Private Sub cboPage_Click()
    Refresh�����嵥 lvwPati.SelectedItem, Val(cboPage.ItemData(cboPage.ListIndex))
End Sub

Private Sub cbo����_Click()
    If cbo����.ListIndex = -1 Then Exit Sub
    ReFresh������Ϣ
'    If lvwPati.ListItems.Count > 0 Then
'        lvwPati_ItemClick lvwPati.ListItems(1)
'    End If
End Sub

Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
     Form_Resize
End Sub

Private Sub lvwALLCLear(ByVal blnCheck As Boolean)
    Dim itm As ListItem
    For Each itm In lvwPati.ListItems
        itm.Checked = blnCheck
    Next
End Sub

Private Sub Form_Activate()
    If cbo����.ListCount = 0 Then
        MsgBox "û�п�˾ְ�Ĳ���(δ��ʼ��Ȩ�޲��߱�)", vbExclamation, "��ʾ"
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    Dim strStartTime  As String
    Dim strEndTime As String
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    RestoreWinState Me, App.ProductName
    
    mnuViewQuitFee.Checked = zlDatabase.GetPara("��ʾ�˷�", glngSys, mlngModul) = "1"
    mnuViewZero.Checked = zlDatabase.GetPara("��ʾ�����", glngSys, mlngModul) = "1"
    mstr����ʱ�� = IIf(zlDatabase.GetPara("����ʱ��", glngSys, mlngModul) = "1", "����ʱ��", "�Ǽ�ʱ��") 'ע���ֵΪ1��ʾ������ʱ��
    
    If InStr(mstrPrivs, ";��������;") = 0 Then
        mnuViewQuitFee.Enabled = False
        mnuViewZero.Enabled = False
    End If
    
    
    mbyt���˲���ģʽ = IIf(zlDatabase.GetPara("���˲���ģʽ", glngSys, mlngModul, "0") = "1", 1, 0)
    
    strEndTime = zlDatabase.GetPara("����ʱ��", glngSys, mlngModul, "23:59:59")
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys, mlngModul, 0))
    If lngTmp > 7 Then lngTmp = 7
    mdtMax = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("��ʼʱ��", glngSys, mlngModul, "00:00:00")
    lngTmp = Val(zlDatabase.GetPara("��ʼ���", glngSys, mlngModul, 0))
    If lngTmp > 7 Then lngTmp = 7
    mdtMin = CDate(Format(mdtMax - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    mbln��ҽ������ = zlDatabase.GetPara("��ҽ������", glngSys, mlngModul, "1") = "1"
    mblnҽ������ = zlDatabase.GetPara("ҽ������", glngSys, mlngModul, "1") = "1"
    
        
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1141")
    
    If InStr(";" & mstrPrivs, ";��Ժ���˲�ѯ;") = 0 Then
        mbln��Ժ���� = True
        mbln��Ժ���� = False
    Else
        mbln��Ժ���� = zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "1") = "1"
        mbln��Ժ���� = zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "1") = "1"
    End If
    
    mblnPrint = True
    If InStr(";" & mstrPrivs, ";�嵥��ӡ;") = 0 Then '�ж��嵥��ӡȨ��
        mblnPrint = False
    End If
    
    
    txtDate.Text = "���ڣ�" & Format(mdtMin, "yyyy��MM��DD�� hh:mm:ss") & "��" & Format(mdtMax, "yyyy��MM��DD�� hh:mm:ss")
    txtDate.Tag = txtDate.Text
    
    
    cbo����.Clear
    If InStr(";" & mstrPrivs, ";���в���;") > 0 Then cbo����.AddItem "���в���"
    Set rsTmp = GetUnit(InStr(mstrPrivs, ";���в���;") = 0, "1,2,3", "����")
    With rsTmp
        Do While Not .EOF
            cbo����.AddItem !����
            cbo����.ItemData(cbo����.NewIndex) = !ID
            If !ID = UserInfo.����ID Then cbo����.ListIndex = cbo����.NewIndex
            .MoveNext
        Loop
        If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    Dim intHeightTbr As Integer, intHeightStb
   
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    intHeightTbr = IIf(cbrThis.Visible, cbrThis.Height, 0)
    intHeightStb = IIf(sta״̬.Visible, sta״̬.Height, 0)
    
    pic�ָ�.Top = 0
    pic�ָ�.Height = ScaleHeight
    If pic�ָ�.Left < 1000 Then pic�ָ�.Left = 1000
    If ScaleWidth - pic�ָ�.Left < 1000 Then pic�ָ�.Left = ScaleWidth - 1000
    
    With lvwPati
        .Top = ScaleTop + intHeightTbr
        .Height = ScaleHeight - intHeightStb - .Top
        .Left = ScaleLeft
        .Width = pic�ָ�.Left - ScaleLeft
    End With
    
    With cboPage
        .Left = pic�ָ�.Left + pic�ָ�.Width
        .Top = ScaleTop + intHeightTbr + 15
    End With
    
    With txtDate
        .Top = ScaleTop + intHeightTbr + 45
        .Left = cboPage.Left + cboPage.Width + 120
        .Width = ScaleWidth - .Left
    End With
    
    With grdList
        .Top = txtDate.Top + txtDate.Height
        .Height = ScaleHeight - intHeightStb - .Top
        .Left = pic�ָ�.Left + pic�ָ�.Width
        .Width = ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmDailyListAsk
        
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwPati.Sorted = True
    With lvwPati
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
    lvwPati.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lvwPati.SelectedItem = Item
    LoadסԺ���� Val(Mid(Item.Key, 2, InStr(Mid(Item.Key, 2), "_") - 1)), Val(Mid(Item.Key, InStr(Mid(Item.Key, 2), "_") + 2))
    Refresh�����嵥 lvwPati.SelectedItem, Val(Mid(Item.Key, InStr(Mid(Item.Key, 2), "_") + 2))
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lvwPati.ListItems.Count = 0 Then Exit Sub
            PopupMenu mnuPop, 2
    End If
End Sub

Private Sub mnuExcel_Click()
    'GrdPrint 1
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = 11
    Call PrintContent(3, Split(lvwPati.SelectedItem.Key, "_")(1))
    Screen.MousePointer = 0
End Sub

Private Sub GrdMuchPrint(ByVal Item As ListItem)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������blnIsPreview false��ʾԤ��
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim blnMuch As Boolean
    Dim i As Long
    Dim old����id As Long
    Dim blnNext As Boolean
    objPrint.Title.Text = GetUnitName & "һ���嵥"
    Set objRow = New zlTabAppRow
    objRow.Add "סԺ�ţ�" & Item.ListSubItems(1).Text & _
        "      ��  ����" & Item.Text & _
        "        �Ա�" & Item.ListSubItems(3).Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��Ժ�գ�" & Item.ListSubItems(5).Text & _
        "  ��Ժ�գ�" & Item.ListSubItems(6).Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add ""
    objRow.Add txtDate
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡʱ��:" & Format(zlDatabase.Currentdate, "yyyy��MM��DD�� HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = grdList
     objPrint.PageFooter = 2
    zlPrintOrView1Grd objPrint, 1
    Set objPrint = Nothing
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    frmDailyListAsk.mlngModul = mlngModul
    frmDailyListAsk.mstrPrivs = mstrPrivs
    frmDailyListAsk.Show 1, Me
    If Not frmDailyListAsk.mblnAskOk Then
        Unload frmDailyListAsk
        Exit Sub
    End If
    
    If frmDailyListAsk.mblnDateMoved Then
        MsgBox "��ǰѡ���ʱ�䷶Χ�ڵķ��ÿ���λ���������ݱ�,���¹��ܽ�������:" & vbCrLf & _
            "��ӡԤ������ӡ�������Excel����ӡ��ѡ����." & vbCrLf & vbCrLf & _
            "��Ҫ������Щ����,�뾡����ѡ����������������ڻ���ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
        Me.mnuFilePrint.Enabled = False
        Me.mnuFilePrintView.Enabled = False
        Me.mnuExcel.Enabled = False
        tbrThis.Buttons(1).Enabled = False
        tbrThis.Buttons(2).Enabled = False
    Else
        Me.mnuFilePrint.Enabled = True
        Me.mnuFilePrintView.Enabled = True
        Me.mnuExcel.Enabled = True
        tbrThis.Buttons(1).Enabled = True
        tbrThis.Buttons(2).Enabled = True
    End If
    
    mstr����ʱ�� = IIf(zlDatabase.GetPara("����ʱ��", glngSys, mlngModul) = "1", "����ʱ��", "�Ǽ�ʱ��") 'ע���ֵΪ1��ʾ������ʱ��
    
    With frmDailyListAsk
        mdtMin = .dtpBegin
        mdtMax = .dtpEnd
        mbln��ҽ������ = .chkPatiType(0).Value = 1
        mblnҽ������ = .chkPatiType(1).Value = 1
        mbln��Ժ���� = .chkInOut(0).Value = 1
        mbln��Ժ���� = .chkInOut(1).Value = 1
        txtDate.Text = "���ڣ�" & Format(mdtMin, "yyyy��MM��DD�� hh:mm:ss") & "��" & Format(mdtMax, "yyyy��MM��DD�� hh:mm:ss")
        txtDate.Tag = txtDate.Text
        mbyt���˲���ģʽ = IIf(.optUnit(0).Value = True, 0, 1)
    End With
    
    Call ReFresh������Ϣ
    
    If lvwPati.SelectedItem Is Nothing Then zlCommFun.StopFlash: Exit Sub
    If lvwPati.ListItems.Count = 0 Then zlCommFun.StopFlash: Exit Sub
    
    Call Refresh�����嵥(lvwPati.SelectedItem)
End Sub

Private Sub mnuFilePrint_Click()
    Dim Item As ListItem
    Dim newPatiId As Long
    Dim blnNOSelect As Boolean
    Dim intPreIdx As Integer, lngCount As Long
    Dim lng��ҳID As Long
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    intPreIdx = lvwPati.SelectedItem.Index
    For Each Item In lvwPati.ListItems
        If Item.Checked Then lngCount = lngCount + 1
    Next
    
    If lngCount > 0 Then
        If MsgBox("��ȷ��Ҫ��ӡ��ѡ���˵�һ�շ����嵥!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    blnNOSelect = True
    grdList.Redraw = False
    For Each Item In lvwPati.ListItems
        If Item.Checked Or (lngCount = 0 And Item Is lvwPati.SelectedItem) Then
            blnNOSelect = False
            
            Item.Selected = True
            Item.EnsureVisible
            Me.Refresh
            lng��ҳID = 0
            If (lngCount = 1 And lvwPati.SelectedItem.Key = Item.Key) Or (lngCount = 0 And Item Is lvwPati.SelectedItem) Then
                If cboPage.ListIndex >= 0 Then
                    lng��ҳID = cboPage.ItemData(cboPage.ListIndex)
                End If
            End If
            Call PrintContent(2, Split(Item.Key, "_")(1), lng��ҳID)
        End If
    Next
    If blnNOSelect Then MsgBox "û��ѡ��Ҫ��ӡ�嵥�Ĳ��ˣ�", vbInformation, gstrSysName
    grdList.Redraw = True
    
    lvwPati.ListItems(intPreIdx).Selected = True
    lvwPati.SelectedItem.EnsureVisible
    
    Screen.MousePointer = 0
End Sub

Private Sub mnuFilePrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me)
End Sub

Private Sub mnuFilePrintView_Click()
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    Screen.MousePointer = 11
    Call PrintContent(1, Split(lvwPati.SelectedItem.Key, "_")(1), Val(cboPage.ItemData(cboPage.ListIndex)))
    Screen.MousePointer = 0
End Sub

Private Sub PrintContent(ByVal bytMode As Byte, ByVal str����ID As String, Optional lng��ҳID As Long = 0)
    ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "����ID=" & str����ID, _
        "��ʼʱ��=" & Format(mdtMin, "yyyy-MM-dd HH:mm:ss"), _
        "����ʱ��=" & Format(mdtMax, "yyyy-MM-dd HH:mm:ss"), _
        "��ʾ�˷�=" & IIf(mnuViewQuitFee.Checked, "1", "0"), _
        "��ʾ�����=" & IIf(mnuViewZero.Checked, "1", "0"), _
        "���˲���=" & cbo����.ItemData(cbo����.ListIndex), _
        "��ҳID=" & lng��ҳID, _
        "����ʱ��=" & mstr����ʱ��, bytMode
End Sub

Private Sub mnuHelp_About_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub
Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuPopDisp_Click()
   If mnuPopDisp.Checked = False Then
        mnuPopDisp.Checked = True
        lvwPati.View = lvwReport
    Else
        mnuPopDisp.Checked = False
        lvwPati.View = lvwIcon
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����ID As Long, lng��ҳID As Long, strסԺ�� As String, lng����ID As Long
    
    If cbo����.ListIndex <> -1 Then
        lng����ID = cbo����.ItemData(cbo����.ListIndex)
    End If
    
    If Not lvwPati.SelectedItem Is Nothing Then
        lng����ID = Val(Mid(lvwPati.SelectedItem.Key, 2, InStr(Mid(lvwPati.SelectedItem.Key, 2), "_") - 1))
        lng��ҳID = Val(Mid(lvwPati.SelectedItem.Key, InStr(Mid(lvwPati.SelectedItem.Key, 2), "_") + 2))
        strסԺ�� = lvwPati.SelectedItem.SubItems(1)
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, "����=" & lng����ID, _
            "��ʼʱ��=" & Format(mdtMin, "yyyy-MM-dd HH:mm:ss "), _
            "����ʱ��=" & Format(mdtMax, "yyyy-MM-dd HH:mm:ss "), "סԺ��=" & strסԺ��)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����=" & lng����ID)
    End If
End Sub

Private Sub mnuViewALLClear_Click()
    lvwALLCLear False
End Sub

Private Sub mnuViewALLSele_Click()
    lvwALLCLear True
End Sub

Private Sub mnuViewFind_Click()
    Dim strBed As String
    
    Load frmPatiFeeFind
    
    With frmPatiFeeFind
        .Show 1, Me
        If Not gblnOK Then Unload frmPatiFeeFind: Exit Sub
                
        strBed = .txtBed.Text
        If mintBedLen - Len(strBed) > 0 Then
            strBed = String(mintBedLen - Len(strBed), " ") & strBed
        End If
        
        mrsPati.Filter = 0
        mrsPati.Filter = "סԺ��=" & Val(.txtסԺ��) & _
            IIf(Trim(.txtBed) = "", "", " Or ����='" & strBed & "'") & _
            IIf(Trim(.txt����) = "", "", " Or ���� Like '" & gstrLike & Trim(.txt����) & "%'")
    End With
    Unload frmPatiFeeFind
    
    If mrsPati.RecordCount = 0 Then
        MsgBox "�޴���Ϣ�Ĳ���!", vbInformation, gstrSysName
        mrsPati.Filter = 0: Exit Sub
    End If
    mrsPati.MoveFirst
    lvwPati.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID).Selected = True
    lvwPati.SelectedItem.EnsureVisible
    Call lvwPati_ItemClick(lvwPati.SelectedItem)
End Sub

Private Sub mnuViewFindNext_Click()
    On Error Resume Next
    If mrsPati Is Nothing Then Exit Sub
    If mrsPati.RecordCount = 0 Or mrsPati.RecordCount = 1 Then Exit Sub
    If mrsPati.EOF Then
        mrsPati.MoveFirst
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then
            mrsPati.MoveFirst
        End If
    End If
    lvwPati.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID).Selected = True
    lvwPati.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID).EnsureVisible
    Call lvwPati_ItemClick(lvwPati.SelectedItem)
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Long
    For i = 0 To 2
        mnuViewFontSize(i).Checked = False
    Next
        mnuViewFontSize(Index).Checked = True
    Select Case Index
    Case 0
        lvwPati.Font.Size = 9
        grdList.Font.Size = 9
        grdList.FontFixed = 9
    Case 1
        lvwPati.Font.Size = 11
        grdList.Font.Size = 11
        grdList.FontFixed = 11
    Case 2
        lvwPati.Font.Size = 12
        grdList.Font.Size = 12
        grdList.FontFixed = 12
    End Select
    Form_Resize
End Sub

Private Sub mnuViewQuitFee_Click()
    mnuViewQuitFee.Checked = Not mnuViewQuitFee.Checked
    zlDatabase.SetPara "��ʾ�˷�", IIf(mnuViewQuitFee.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    lvwPati_ItemClick lvwPati.SelectedItem
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta״̬.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub
Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    mnuViewToolbarStand.Checked = Not mnuViewToolbarStand.Checked
    mnuViewToolbarText.Enabled = mnuViewToolbarStand.Checked
    cbrThis.Visible = mnuViewToolbarStand.Checked
    If mnuViewToolbarText.Checked Then
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    cbrThis.Bands(1).minHeight = tbrThis.Height
    cbrThis.Refresh
    Form_Resize
End Sub

Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    mnuViewToolbarText.Checked = Not mnuViewToolbarText.Checked
    If mnuViewToolbarText.Checked Then
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    cbrThis.Bands(1).minHeight = tbrThis.Height
    cbrThis.Refresh
    Form_Resize
End Sub

Private Sub mnuViewZero_Click()
    mnuViewZero.Checked = Not mnuViewZero.Checked
    zlDatabase.SetPara "��ʾ�����", IIf(mnuViewZero.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    lvwPati_ItemClick lvwPati.SelectedItem
End Sub

Private Sub pic�ָ�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        pic�ָ�.Left = pic�ָ�.Left + X
        Form_Resize
        Me.Refresh
    End If
End Sub

Private Sub sta״̬_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "������ɫ" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "Ԥ��"
            mnuFilePrintView_Click
        Case "��ӡ"
            mnuFilePrint_Click
        Case "����"
            mnuFileOpen_Click
        Case "����"
            mnuViewFind_Click
        Case "����"
             PopupMenu mnuViewFont
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
           mnuFileExit_Click
        End Select
    End With
  
End Sub

Private Sub LoadסԺ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim strSql As String, rsPage As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Distinct ��ҳID From ������ҳ Where ����ID = [1] And �������� = 0 Order By ��ҳID Desc"
    Set rsPage = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    cboPage.Clear
    cboPage.AddItem "����סԺ"
    cboPage.ItemData(cboPage.NewIndex) = 0
    Do While Not rsPage.EOF
        cboPage.AddItem "��" & Val(NVL(rsPage!��ҳID)) & "��סԺ"
        cboPage.ItemData(cboPage.NewIndex) = Val(NVL(rsPage!��ҳID))
        If lng��ҳID = Val(NVL(rsPage!��ҳID)) Then cboPage.ListIndex = cboPage.NewIndex
        rsPage.MoveNext
    Loop
    If cboPage.ListIndex < 0 Then cboPage.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub Refresh�����嵥(Item As ListItem, Optional ByVal lngPageID As Long = 0)
    Dim rsTmp As ADODB.Recordset
    Dim arrFields As Variant, strSql As String
    Dim lngRow As Long, lngCol As Integer
    Dim strTmp As String, i As Long
    
    Dim lng����ID As Long, lng��ҳID As Long, lng����ID As Long, lngInsure As Long
    
    On Error GoTo errHandle
    
    lngInsure = Val("" & Item.Tag)
    lng����ID = Val(Mid(Item.Key, 2, InStr(Mid(Item.Key, 2), "_") - 1))

    lng��ҳID = lngPageID

    If mbyt���˲���ģʽ = 0 And cbo����.ListIndex <> -1 Then '���÷����Ĳ���
        lng����ID = cbo����.ItemData(cbo����.ListIndex)
    End If
    
    '�������˷�:������ܵ��շ�ϸĿ��,����ÿ���˷ѵ�����,���
    strSql = _
    " Select Mod(��¼����,10) as ��¼����,NO,Nvl(�۸񸸺�,���) as ���,�շ�ϸĿID," & _
    "       ���㵥λ,Avg(Nvl(����,1)) as ����,Avg(����) as ����," & _
    "       Sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��,����ʱ��,�������� " & _
    " From " & IIf(frmDailyListAsk.mblnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼") & _
    " Where ��¼״̬<>0 And ���ʷ���=1 And ����ID=[1] " & IIf(lng��ҳID = 0, "", " And ��ҳID=[2] ") & _
            IIf(lng����ID = 0, "", " And ���˲���ID=[3]") & _
    "       And " & mstr����ʱ�� & " Between [4] And [5]" & _
    " Group by Mod(��¼����,10),NO,��¼״̬,Nvl(�۸񸸺�,���),�շ�ϸĿID,���㵥λ,ִ��״̬,����ʱ��,�������� "
    
    '�����˷�:������ܵ��շ�ϸĿ�е�ʣ������,���
    If Not mnuViewQuitFee.Checked Then
            strSql = _
            " Select ��¼����,NO,���,�շ�ϸĿID,���㵥λ," & _
            " Sum(����) as ����,Sum(����) as ����," & _
            " Sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��,����ʱ��,��������" & _
            " From (" & strSql & ")" & _
            " Group by ��¼����,NO,���,�շ�ϸĿID,���㵥λ,����ʱ��,��������"
    End If
    
    '�Ƿ���ʾ�����
    If mnuViewZero.Checked Then
        strSql = strSql & " Having Nvl(Sum(Ӧ�ս��),0)<>0"
    Else
        strSql = strSql & " Having Nvl(Sum(ʵ�ս��),0)<>0"
    End If
    
    strSql = _
        " Select To_Char(L.����ʱ��,'YYYY-MM-DD') as ����,L.NO as ���ݺ�," & _
        " Nvl(X.����,I.����)||' '||I.���||'   '||LTrim(To_Char(L.����,'9999990.00000'))||L.���㵥λ||Decode(I.���,'7','��'||L.����||'��',NULL) as ��Ŀ," & _
        " LTrim(To_Char(L.ʵ�ս��,'9999999" & gstrDec & "')) as ���,NVL(L.��������,I.��������) as ��������,N.���� ҽ������" & _
        " From �շ���ĿĿ¼ I,(" & strSql & ") L,�շ���Ŀ���� X,����֧����Ŀ M,����֧������ N" & _
        " Where I.ID=L.�շ�ϸĿID And I.ID=X.�շ�ϸĿID(+) And X.����(+)=1 And X.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        " And I.ID=M.�շ�ϸĿID(+) And M.����(+)=[6] And M.����ID=N.ID(+)" & vbNewLine & _
        " Order by ����,���ݺ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID, lng����ID, mdtMin, mdtMax, lngInsure)
    
    grdList.Redraw = False
    grdList.Clear
    grdList.Rows = 2
    If Not rsTmp.EOF Then
        Set grdList.Recordset = rsTmp
        For i = 0 To grdList.Cols - 1
            grdList.ColAlignmentFixed(i) = 4
            Select Case grdList.TextMatrix(0, i)
                Case "���"
                    grdList.ColAlignment(i) = 7
                    strTmp = strTmp & "," & i
                Case Else
                    grdList.ColAlignment(i) = 1
            End Select
        Next
        If strTmp <> "" Then
            grdList.Rows = grdList.Rows + 1
            arrFields = Split(Mid(strTmp, 2), ",")
            For i = 0 To grdList.Rows - 1
                If i <> grdList.Rows - 1 Then
                     For lngCol = 0 To UBound(arrFields)
                         grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol)) = Val(grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol))) + Val(grdList.TextMatrix(i, arrFields(lngCol)))
                         grdList.TextMatrix(grdList.Rows - 1, 0) = "�ϼ�"
                     Next
                End If
                Call RefreshGridColWidth(grdList, i)
            Next
            For lngCol = 0 To UBound(arrFields)
                grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol)) = Format(Val(grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol))), "####" & gstrDec & ";-####" & gstrDec & "; ;")
            Next
        End If
        If lngInsure = 0 Then
            grdList.ColWidth(MshGetColNum(grdList, "ҽ������")) = 0
        Else
            grdList.ColWidth(MshGetColNum(grdList, "ҽ������")) = grdList.ColWidth(MshGetColNum(grdList, "��������"))
        End If
    End If
    grdList.Row = 1: grdList.Col = 0
    grdList.ColSel = grdList.Cols - 1
    
    grdList.Redraw = True
    
    strTmp = GetPatientDue(lng����ID)
    If Val(strTmp) <> 0 Then
        txtDate.Text = txtDate.Tag & "" & "��Ӧ�տ�:" & Format(strTmp, "0.00")
    Else
        txtDate.Text = txtDate.Tag
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ReFresh������Ϣ()
    Dim objItem As ListItem
    Dim strSql As String, i As Integer, lng����ID As Long
    
    Call zlCommFun.ShowFlash("����ͳ������,���Ժ� ...", Me)
    DoEvents

    On Error GoTo errHandle
    
    grdList.Clear
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    mintBedLen = GetMaxBedLen(lng����ID)


    strSql = " Where " & mstr����ʱ�� & " Between [1] And [2] And ��¼״̬ IN(1,2,3) And ���ʷ���=1 And ��¼���� In (2,3,5) "
        
    If mbyt���˲���ģʽ = 0 Then
        If lng����ID > 0 Then strSql = strSql & " And ���˲���id+0=[3]"
    End If
            
    If frmDailyListAsk.mblnDateMoved Then
       strSql = "Select Distinct ����id From (Select ����id From סԺ���ü�¼ " & strSql & _
                " Union All Select ����id From HסԺ���ü�¼ " & strSql & ")"
    Else
       strSql = "Select Distinct ����id From סԺ���ü�¼ " & strSql
    End If
    
    strSql = "" & _
    "   Select /*+ rule*/ I.����id,P.��ҳid,nvl(P.����,I.����) as ����,I.סԺ��,LPAD(P.��Ժ����," & mintBedLen & ",' ') as ����," & _
    "           nvl(P.�Ա�,I.�Ա�) as �Ա�,P.��Ժ����,P.��Ժ����,P.��������,P.�ѱ�,X.���� as ����,P.��������,P.����" & _
    "   From ���ű� X,������Ϣ I,������ҳ P,(" & strSql & " ) L" & _
    "   Where I.����id=P.����id and P.��ҳid=I.��ҳid and P.����id=L.����id And P.��ǰ����ID=X.ID" & _
    "           And (X.վ��=[4] or X.վ�� is NULL)"
        
    If mbyt���˲���ģʽ = 1 Then
        If lng����ID > 0 Then strSql = strSql & " And P.��ǰ����ID+0=[3]"
    End If
    
    
    '��Ժ���Ժ����
    If mbln��Ժ���� And mbln��Ժ���� Then
    ElseIf mbln��Ժ���� Then
        strSql = strSql & " And P.��Ժ���� is NULL"
    ElseIf mbln��Ժ���� Then
        strSql = strSql & " And P.��Ժ���� is Not NULL"
    End If
    
    'ҽ������ͨ����
    If mbln��ҽ������ And mblnҽ������ Then
    ElseIf mbln��ҽ������ Then
        strSql = strSql & " And P.���� is NULL"
    ElseIf mblnҽ������ Then
        strSql = strSql & " And P.���� is Not NULL"
    End If
    
    If cbo����.ItemData(cbo����.ListIndex) = 0 Then
        strSql = strSql & " Order BY ����,LPAD(����,10,' ')"
    Else
        strSql = strSql & " Order BY LPAD(����,10,' ')"
    End If
    
    
    Set mrsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtMin, mdtMax, lng����ID, gstrNodeNo)
    With mrsPati
         If .RecordCount <> 0 And mblnPrint Then
            If Not frmDailyListAsk.mblnDateMoved Then
             tbrThis.Buttons.Item(1).Enabled = True
             tbrThis.Buttons.Item(2).Enabled = True
             mnuFilePrint.Enabled = True
             mnuFilePrintView.Enabled = True
             mnuExcel.Enabled = True
            End If
        Else
             tbrThis.Buttons.Item(1).Enabled = False
             tbrThis.Buttons.Item(2).Enabled = False
             mnuFilePrint.Enabled = False
             mnuFilePrintView.Enabled = False
             mnuExcel.Enabled = False
         End If
        
        .Filter = 0
        lvwPati.ListItems.Clear
        Do While Not .EOF
            If IIf(IsNull(!��������), 0, !��������) = 0 Then
                Set objItem = lvwPati.ListItems.Add(, "_" & !����ID & "_" & !��ҳID, !����, 1, 1)
            Else
                Set objItem = lvwPati.ListItems.Add(, "_" & !����ID & "_" & !��ҳID, !����, 2, 2)
            End If
            objItem.ListSubItems.Add , , IIf(IsNull(!סԺ��), "", !סԺ��)
            objItem.ListSubItems.Add , , IIf(IsNull(!����), "", !����)
            objItem.ListSubItems.Add , , IIf(IsNull(!�ѱ�), "", !�ѱ�)
            objItem.ListSubItems.Add , , IIf(IsNull(!�Ա�), "", !�Ա�)
            objItem.ListSubItems.Add , , Format(!��Ժ����, "yyyy-MM-DD")
            objItem.ListSubItems.Add , , Format(IIf(IsNull(!��Ժ����), Empty, !��Ժ����), "yyyy-MM-DD")
            objItem.ListSubItems.Add , , IIf(IsNull(!����), "", !����)
            objItem.ListSubItems.Add , , IIf(IsNull(!��������), "", !��������)
            objItem.Tag = Val("" & !����)
        
            objItem.ForeColor = zlDatabase.GetPatiColor(NVL(!��������))
            For i = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(i).ForeColor = zlDatabase.GetPatiColor(NVL(!��������))
            Next
            .MoveNext
        Loop
    End With
    sta״̬.Panels(2).Text = "��" & lvwPati.ListItems.Count & "��"
    If mrsPati.RecordCount = 0 Then Call RefreshListStru
    If Not lvwPati.SelectedItem Is Nothing Then
         lvwPati.SelectedItem.Selected = False
         Set lvwPati.SelectedItem = Nothing:
         Call RefreshListStru
    End If
    Call zlCommFun.StopFlash
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshListStru()
    '--------------------------------------------------------------
    '���ܣ���ȡ���˷����嵥�ı�ͷ�ṹ
    '������
    '���أ�
    '--------------------------------------------------------------
   
   Dim intRow As Long
   Dim intCol As Long
    '0  ��ʾ�����嵥
   
   With grdList
        .Redraw = False
        For intRow = 0 To .Cols - 1
            .MergeCol(intRow) = False
        Next
        For intRow = 1 To .Rows - 1
            .RowData(intRow) = 0
            .MergeRow(intRow) = False
        Next
        .Clear
        .Rows = 2
        .FixedRows = 1
        .RowHeight(0) = TextHeight("��") * 2
        .MergeCells = flexMergeRestrictRows
        .Cols = 11
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1
        
        .ColWidth(0) = 1400
        .ColWidth(1) = 800
        .ColWidth(2) = 1600
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
        .ColWidth(10) = 600
        
        .MergeCol(0) = False
        .MergeCol(1) = False
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        .MergeCol(5) = False
        .MergeCol(6) = False
        .MergeCol(7) = False
        .MergeCol(8) = False
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "���ݺ�"
        .TextMatrix(0, 2) = "ժҪ"
        .TextMatrix(0, 3) = "��������"
        .TextMatrix(0, 4) = "�վ���Ŀ"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "Ӧ�ս��"
        .TextMatrix(0, 7) = "ʵ�ս��"
        .TextMatrix(0, 8) = "����"
        .TextMatrix(0, 9) = "����Ա"
        .TextMatrix(0, 10) = "����Ա"
        
        For intCol = 0 To .Cols - 1
          .ColAlignmentFixed(intCol) = 4
        Next
        .Redraw = True
     End With
    Call RefreshGridColWidth(grdList, 0)
End Sub

Private Sub GrdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������
    '     blnIsPreview: 0��ʾԤ�� 1��ʾ�����EXCEL ������ʾ��ӡ
    '���أ�
    '---------------------------------------------------
    '0 ��ʾ������ϸ 1��ʾ���ջ����嵥 2��ʾԤ����ϸ�嵥,3������ϸ,4 δ�����
    
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    objPrint.Title.Text = GetUnitName & "����һ���嵥"
    Set objRow = New zlTabAppRow
    objRow.Add "סԺ�ţ�" & lvwPati.SelectedItem.ListSubItems(1).Text & _
        "      ��  ����" & lvwPati.SelectedItem.Text & _
        "        �Ա�" & lvwPati.SelectedItem.ListSubItems(3).Text & _
        ""
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��Ժ�գ�" & lvwPati.SelectedItem.ListSubItems(5).Text
    objRow.Add "  ��Ժ�գ�" & lvwPati.SelectedItem.ListSubItems(6).Text
    objPrint.UnderAppRows.Add objRow
     
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add ""
    objRow.Add txtDate
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡʱ��:" & Format(zlDatabase.Currentdate, "yyyy��MM��DD��")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = grdList
    objPrint.PageFooter = 2
    If blnIsPreview = 0 Then
        zlPrintOrView1Grd objPrint, 2
    Else
        If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
        Else
            Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
            End Select
        End If
    End If
    Set objPrint = Nothing
    
End Sub

Private Sub RefreshGridColWidth(ByVal objGrid As Object, lngRow As Long)
    Dim lngWidth As Long, lngCol As Long
    
    For lngCol = 0 To objGrid.Cols - 1
        lngWidth = Me.TextWidth(objGrid.TextMatrix(lngRow, lngCol) & "��")
        If objGrid.ColWidth(lngCol) <> 0 Then
            If objGrid.ColWidth(lngCol) < lngWidth Or lngRow = 0 Then
                objGrid.ColWidth(lngCol) = lngWidth
            End If
        End If
    Next
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
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

