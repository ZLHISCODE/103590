VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDrugQualityList 
   Caption         =   "ҩƷ��������"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9750
   Icon            =   "frmDrugQualityList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9750
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "�ⷿ"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   3615
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   6825
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2835
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
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
               Key             =   "Verify"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Clear"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugQualityList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugQualityList.frx":0464
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
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
      Left            =   2280
      Top             =   600
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
            Picture         =   "frmDrugQualityList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":0F14
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1130
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":134A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1564
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":177E
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1998
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2092
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":22AC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2622
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   1545
      Top             =   600
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
            Picture         =   "frmDrugQualityList.frx":283E
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2A5A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2C76
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2E90
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":30AC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":32C6
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":34E0
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3BDA
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3DF4
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3F4E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":416A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1965
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   6255
      _cx             =   11033
      _cy             =   3466
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQualityList.frx":4386
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   3330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "����(&V)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
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
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmDrugQualityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mblnBootUp As Boolean
Private mstrPrivs As String
Private mblnViewCost As Boolean         '�鿴�ɱ��� true-����鿴 flase-������鿴
Private Const MStrCaption As String = "ҩƷ��������"

Private mbln��ҩ�ⷿ As Boolean

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private Type Type_SQLCondition
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng��Ӧ�� As Long
    str������ As String
    str����� As String
End Type

Private SQLCondition As Type_SQLCondition

Private mlng�ⷿid As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4
'�������������
Private Function CheckDepend() As Boolean
    Dim rsDepend As New Recordset
    
    CheckDepend = False
    On Error GoTo errHandle
    If InStr(mstrPrivs, "���пⷿ") = 0 Then
        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a, ������Ա d " _
                & "Where (a.վ�� = [1] Or a.վ�� is Null) And c.�������� = b.���� " _
                & "  AND Instr('HIJKLMN',b.����,1) > 0 " _
                & "  AND a.id = c.����id AND a.id=d.����id and d.��Աid=[2] " _
                & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                & "Where (a.վ�� = [1] Or a.վ�� is Null) And c.�������� = b.���� " _
                & "  AND Instr('HIJKLMN',b.����,1) > 0 " _
                & "  AND a.id = c.����id " _
                & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    End If
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.�û�ID)
    
    If rsDepend.EOF Then
        MsgBox "����Ӧ������һ������ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get���ⵥ�ݺ�(ByVal lngBillId As Long)
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ���ⵥNo From ҩƷ������¼ Where Id = [1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���ⵥ�ݺ�", lngBillId)
    
    If Not rstemp.EOF Then
        Get���ⵥ�ݺ� = IIf(IsNull(rstemp!���ⵥno), "", rstemp!���ⵥno)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboStock_Click()
    If mlng�ⷿid <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng�ⷿid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '��֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        If mblnBootUp Then mnuViewRefresh_Click
    End If
End Sub

Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsList As ADODB.Recordset
    Dim strUnit As String
    Dim str��װϵ�� As String
    Dim str����ϵ�� As String
    Dim strSqlҩ�� As String
    Dim n As Integer
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    Call FS.ShowFlash("��������ҩƷ���������¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    vsfList.Redraw = False
    
    mbln��ҩ�ⷿ = False
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsList!��������
        rsList.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then mbln��ҩ�ⷿ = True
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnit = "F.���㵥λ"
            str��װϵ�� = "1 as ����ϵ�� "
            str����ϵ�� = "1 "
        Case mconint���ﵥλ
            strUnit = "B.���ﵥλ"
            str��װϵ�� = "B.�����װ as ����ϵ�� "
            str����ϵ�� = "B.�����װ "
        Case mconintסԺ��λ
            strUnit = "B.סԺ��λ"
            str��װϵ�� = "B.סԺ��װ as ����ϵ�� "
            str����ϵ�� = "B.סԺ��װ "
        Case mconintҩ�ⵥλ
            strUnit = "B.ҩ�ⵥλ"
            str��װϵ�� = "B.ҩ���װ as ����ϵ�� "
            str����ϵ�� = "B.ҩ���װ "
    End Select
        
    If gintҩƷ������ʾ = 0 Then
        strSqlҩ�� = ",('['||F.����||']'||F.����) AS ҩƷ��Ϣ"
    ElseIf gintҩƷ������ʾ = 1 Then
        strSqlҩ�� = ",('['||F.����||']'||NVL(D.����,F.����)) AS ҩƷ��Ϣ"
    Else
        strSqlҩ�� = ",('['||F.����||']'||F.����) AS ҩƷ��Ϣ,D.���� As ��Ʒ��"
    End If
        
    gstrSQL = "SELECT DISTINCT A.ID" & strSqlҩ�� & _
        ",A.ҩƷID,A.����,A.��ҩ��λID,F.���,A.���� as ������," & IIf(mbln��ҩ�ⷿ, "B.ԭ����,", "") & "A.����," & _
        " ltrim(to_char(a.�ɱ�����*" & str����ϵ�� & "," & mstrCostFormat & ")) as �ɱ���," & _
        " ltrim(to_char(a.�ɱ����," & mstrMoneyFormat & ")) as �ɱ����," & _
        " ltrim(to_char(a.���۵���*" & str����ϵ�� & "," & mstrPriceFormat & ")) as ���ۼ�," & _
        " ltrim(to_char(a.���۽��," & mstrMoneyFormat & ")) as ���۽��," & _
        strUnit & " AS ��λ,LTRIM(TO_CHAR(A.��������/(" & str����ϵ�� & ")," & mstrNumberFormat & ")) AS ��������," & str��װϵ�� & _
        " ,A.����ԭ��,C.���� AS ��Ӧ��,A.�Ǽ���,TO_CHAR(A.�Ǽ�ʱ��,'YYYY-MM-DD') AS �Ǽ�ʱ��,A.����취,A.������,TO_CHAR(A.����ʱ��,'YYYY-MM-DD') AS ����ʱ�� " & _
        " FROM ҩƷ������¼ A, ҩƷ��� B, �շ���ĿĿ¼ F, �շ���Ŀ���� D, ��Ӧ�� C " & _
        " WHERE A.ҩƷID = B.ҩƷID And B.ҩƷID=F.ID " & _
        " AND F.id = D.�շ�ϸĿID(+) AND A.��ҩ��λID = C.ID(+) And D.����(+)=3 And D.����(+)=1" & _
        " AND SUBSTR(����(+),1,1)='1' " & _
        " AND A.�ⷿID = [9] " & _
        strFind & _
        " ORDER BY �Ǽ�ʱ�� DESC "

    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng��Ӧ��, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        cboStock.ItemData(cboStock.ListIndex))
        
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            .TopRow = 1
            .rows = .rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    vsfList.Redraw = True
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(8 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = flexAlignRightCenter
        .ColAlignment(9 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = flexAlignRightCenter
        .ColAlignment(10 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = flexAlignRightCenter
        .ColAlignment(11 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = flexAlignRightCenter
        .ColAlignment(13 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = flexAlignRightCenter
        
        If mblnBootUp = False Then
'            For intCol = 0 To .Cols - 1
'                .ColWidth(intCol) = 1500
'            Next
            .ColWidth(1) = 2000
            If gintҩƷ������ʾ = 2 Then .ColWidth(2) = 2000
            .ColWidth(5 + IIf(gintҩƷ������ʾ = 2, 1, 0)) = 2000
        End If
        If mblnViewCost = False Then
            .ColWidth(9 + IIf(mbln��ҩ�ⷿ, 1, 0)) = 0 '�ɱ���
            .ColWidth(10 + IIf(mbln��ҩ�ⷿ, 1, 0)) = 0 '�ɱ����
        End If
        
        .ColWidth(0) = 0
        .ColWidth(2 + IIf(gintҩƷ������ʾ = 2, 1, 0)) = 0
        .ColWidth(3 + IIf(gintҩƷ������ʾ = 2, 1, 0)) = 0
        .ColWidth(4 + IIf(gintҩƷ������ʾ = 2, 1, 0)) = 0
        .ColWidth(14 + IIf(gintҩƷ������ʾ = 2, 1, 0) + IIf(mbln��ҩ�ⷿ, 1, 0)) = 0
        
    End With
End Sub

'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա�����
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "�����Ǽ�") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        mnuEditLine1.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        
        tlbTool.Buttons("EditSeparate").Visible = False
        
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "�������") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "�����¼") Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
        
End Sub


Private Sub Form_Load()
    '�ָ�����
    Dim strStart As String
    Dim strEnd As String
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int��ѯ���� As Integer
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    mblnBootUp = False
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    mlng�ⷿid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿid, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
        
    dateCurrentDate = Sys.Currentdate
    int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.����ʱ�� is Null And A.�Ǽ�ʱ�� Between [1] And [2] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    mstrFind = strFind
    
    lblRange.Caption = "��ѯ��Χ:" & Format(dateCurrentDate, "yyyy��MM��dd��") & "��" & Format(dateCurrentDate, "yyyy��MM��dd��")
    GetList (mstrFind)  '�г�����ͷ
    '�ָ����Ի�����
    RestoreWinState Me, App.ProductName, MStrCaption
    '�ָ����Ի����ú�Ȩ�޿��Ƶ��л���Ҫ��һ������
    vsfList.ColWidth(9) = IIf(mblnViewCost = True, 1000, 0) '�ɱ���
    vsfList.ColWidth(10) = IIf(mblnViewCost = True, 1400, 0) '�ɱ����
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With lblRange
        .Top = IIf(staThis.Visible = True, Me.ScaleHeight - staThis.Height - .Height - 100, Me.ScaleHeight - .Height - 100)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = lblRange.Top - .Top - 50
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
   
End Sub


Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    strNo = ""
    '����
    BlnSuccess = frmDrugQualityCard.ShowCard(Me, 1, 0, mstrPrivs)
    
    
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditClear_Click()
    'ɾ��
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim strNo As String
    Dim StrDate As String
    
    Dim Dbl���� As Double
    Dim strsql As String
    Dim rscord As Recordset
    
    On Error GoTo errHandle
    strsql = "select �������� from ҩƷ������¼ where id=[1]"
    Set rscord = zldatabase.OpenSQLRecord(strsql, "mnuEditClear_Click", vsfList.TextMatrix(vsfList.Row, 0))
    If rscord.EOF Then
       MsgBox "��ҩƷ������¼�ѱ�������ɾ�������飡", vbOKOnly, gstrSysName
       Exit Sub
    Else
        Dbl���� = rscord!��������
    End If
    rscord.Close
    
    With vsfList
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪ���ҩƷ��ϢΪ��" & .TextMatrix(.Row, 1) & "����ҩƷ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            strNo = Get���ⵥ�ݺ�(lngBillId)
            
            If Trim(strNo) <> "" Then gcnOracle.BeginTrans
            
            gstrSQL = "zl_ҩƷ��������_delete(" & lngBillId & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            If Trim(strNo) <> "" Then
                StrDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                
                gstrSQL = "ZL_ҩƷ��������_STRIKE("
                '�д�
                gstrSQL = gstrSQL & "1"
                'ԭ��¼״̬
                gstrSQL = gstrSQL & ",1"
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '���
                gstrSQL = gstrSQL & ",1"
                'ҩƷID
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 2 + IIf(gintҩƷ������ʾ = 2, 1, 0)))
                '��������
                'gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 9 + IIf(gintҩƷ������ʾ = 2, 1, 0))) * Val(.TextMatrix(intRow, 10 + IIf(gintҩƷ������ʾ = 2, 1, 0)))
                'gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 13 + IIf(gintҩƷ������ʾ = 2, 1, 0)))
                gstrSQL = gstrSQL & "," & Dbl����
                
                '������
                gstrSQL = gstrSQL & ",'" & UserInfo.�û����� & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & Format(StrDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                
                gcnOracle.CommitTrans
            End If
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    If Trim(strNo) <> "" Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim lngRecordID As Long
    Dim BlnSuccess As Boolean
    
    With vsfList
        lngRecordID = .TextMatrix(.Row, 0)
        BlnSuccess = frmDrugQualityCard.ShowCard(Me, 3, lngRecordID, mstrPrivs)
    
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim rsParrel As New Recordset
    
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ��ҩƷ��ϢΪ��" & .TextMatrix(.Row, 1) & "����ҩƷ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "select nvl(������,'0') from ҩƷ������¼ where id=[1]"
            Set rsParrel = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngBillId)
            
            If rsParrel.EOF Then
                MsgBox "��ҩƷ������¼�ѱ�������ɾ�������飡", vbOKOnly, gstrSysName
                Exit Sub
            ElseIf rsParrel.Fields(0) <> "0" Then
                MsgBox "��ҩƷ������¼�ѱ������˴������飡", vbOKOnly, gstrSysName
                Exit Sub
            End If
            rsParrel.Close
            
            gstrSQL = "zl_ҩƷ��������_delete(" & lngBillId & ")"
        
            If gstrSQL = "" Then Exit Sub
            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim lngRecordID As Long
    
    With vsfList
        lngRecordID = .TextMatrix(.Row, 0)
        frmDrugQualityCard.ShowCard Me, 4, lngRecordID, mstrPrivs
    End With
End Sub

Private Sub mnuEditModify_Click()
    '�޸�
    Dim lngRecordID As Long
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        lngRecordID = .TextMatrix(.Row, 0)
        BlnSuccess = frmDrugQualityCard.ShowCard(Me, 2, lngRecordID, mstrPrivs)
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1331", Me
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    Call mnuFileBillPreview_Click
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    vsfList.Redraw = False
    subPrint 3
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '��������
    Dim dateCurrentDate As Date
    Dim int��ѯ���� As Date
    
    frm��������.���ò��� Me, mstrPrivs, MStrCaption
    
    dateCurrentDate = Sys.Currentdate
    int��ѯ���� = Val(zldatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")

    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    vsfList.Redraw = False
    subPrint 2
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    vsfList.Redraw = False
    subPrint 1
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
'    ReportMan gcnOracle, Me
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid���ⷿ=�ⷿid����Ӧ��=��Ӧ��id����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ��
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    
    str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
    str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
        "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "��Ӧ��=" & IIf(SQLCondition.lng��Ӧ�� = 0, "", SQLCondition.lng��Ӧ��), _
        "��ʼʱ��=" & str��ʼʱ��, _
        "����ʱ��=" & str����ʱ��)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    
    strFind = FrmDrugQualitySearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.lngҩƷ, _
                SQLCondition.lng��Ӧ��, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                cboStock.ItemData(cboStock.ListIndex))
    
    If strFind <> "" Then
        mstrFind = strFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�Ǽ�ʱ�� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ����ʱ�� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�Ǽ�ʱ�� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�Ǽ�ʱ�� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub


Private Sub vsfList_Click()
    With vsfList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditDisplay_Click
End Sub

Private Sub vsfList_EnterCell()
    SetEnable
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditDisplay_Click
    End If
        
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub


Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        
    End Select
    
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    With vsfList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
        Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 3) = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = False
                    tlbTool.Buttons("Clear").Enabled = False
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else    '��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                                    
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = True
                    tlbTool.Buttons("Clear").Enabled = True
                End If
                                    
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = MStrCaption
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
'            mnuEditAddPayment_Click
        Case "Imprest"
'            mnuEditAddImprest_Click
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'�Ե���ͷ������
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With vsfList
        If .rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               mintsort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               mintsort = flexSortStringNoCaseDescending
            End If
            
            mintPreCol = intCol
            .Row = FindRow(vsfList, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'Ѱ����ĳһ����ȵ���
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

