VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPriceQuery 
   BackColor       =   &H8000000A&
   Caption         =   "�շ���Ŀ���Ŀ"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frmPriceQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid msh������Ŀ 
      Height          =   1530
      Left            =   2715
      TabIndex        =   7
      ToolTipText     =   "������Ŀ"
      Top             =   2850
      Width           =   5400
      _cx             =   9525
      _cy             =   2699
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPriceQuery.frx":030A
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
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
   Begin VB.PictureBox picHBar_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   825
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   6075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2250
      Width           =   6075
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   945
      Left            =   3075
      TabIndex        =   5
      Top             =   960
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1667
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2055
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2490
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":03DB
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":06F5
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":084F
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   75
      TabIndex        =   3
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   5640
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":0CA1
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":0EBD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":10D9
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":12F3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":150F
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   4920
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":172B
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1947
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1B63
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1D7D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1F99
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8160
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   635
      SimpleText      =   $"frmPriceQuery.frx":21B5
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPriceQuery.frx":21FC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9313
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
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
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
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
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
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowDynamic 
         Caption         =   "��ʾ�����Ŀ(&D)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R) "
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
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)��"
      End
   End
End
Attribute VB_Name = "frmPriceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Column��Ŀ
    col���� = 0
    col���� = 1
    col��� = 2
    col���� = 3
    col�ۼ۵�λ = 4
    col������Ŀ = 5
    col�۸� = 6
    col���� = 7
    col�Ӱ� = 8
End Enum

Dim strPrivs As String   'ģ��Ȩ��
Dim mblnTradeName As Boolean        '�Ƿ�����Ʒ����ʾ��ҩ
Dim rsTemp As New ADODB.Recordset

Dim mblnLoad As Boolean
Dim msngStartX As Single    '�ƶ�ǰ����λ��
Dim mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Dim mrs��Ŀ As New ADODB.Recordset
Dim msngOldY As Single
Private Type �ۼ۾���
    �շ���ĿС�� As Integer
    ҩƷ��ĿС�� As Integer
    ������ĿС�� As Integer
End Type
Private m�ۼ۾��� As �ۼ۾���

Private mstrVbFormat As String
Private mstrOraFormat As String
Private mlngPreRow As Long

Private Sub Form_Activate()
    If Me.tvwMain_S.Nodes.Count = 0 Then
        MsgBox "û�н�����Ŀ������Ȩ�޲��߱�", vbExclamation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    strPrivs = gstrPrivs
    mblnLoad = True
    RestoreWinState Me, App.ProductName
    mnuViewShowDynamic.Checked = (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewShowDynamic״̬", "False") = "True")
    
    Call initС��λ��
    
    mstrVbFormat = GetFmtString(1, False)
    mstrOraFormat = GetFmtString(1, True)
    
    mrs��Ŀ.CursorLocation = adUseClient
    '�õ���ѯ��ʱ�䷶Χ
    Call InitSum
    
    mblnTradeName = False
'    gstrSQL = "Select nvl(����ֵ,'0') From ϵͳ������ Where ������='����ҩ����Ʒ����ʾ'"
    mblnTradeName = IIf(zlDatabase.GetPara(74, 100, , -1) = 1, True, False)
    
    'װ�����
    With tvwMain_S.Nodes
        .Clear
        If InStr(1, strPrivs, "�շ���Ŀ") <> 0 Then
            .Add , , "K0", "[0]�շ���Ŀ", "R", "R"
            tvwMain_S.Nodes("K0").Sorted = True
            tvwMain_S.Nodes("K0").Tag = "0"
            Call FillTree(0)
        End If
        If InStr(1, strPrivs, "����ҩ") <> 0 Then
            .Add , , "K1", "[1]����ҩ", "R", "R"
            tvwMain_S.Nodes("K1").Sorted = True
            tvwMain_S.Nodes("K1").Tag = 1
            Call FillTree(1)
        End If
        If InStr(1, strPrivs, "�г�ҩ") <> 0 Then
            .Add , , "K2", "[2]�г�ҩ", "R", "R"
            tvwMain_S.Nodes("K2").Sorted = True
            tvwMain_S.Nodes("K2").Tag = 2
            Call FillTree(2)
        End If
        If InStr(1, strPrivs, "�в�ҩ") <> 0 Then
            .Add , , "K3", "[3]�в�ҩ", "R", "R"
            tvwMain_S.Nodes("K3").Sorted = True
            tvwMain_S.Nodes("K3").Tag = 3
            
            Call FillTree(3)
        End If
        If InStr(1, strPrivs, "��������") <> 0 Then
            .Add , , "K7", "[4]��������", "R", "R"
            tvwMain_S.Nodes("K7").Sorted = True
            tvwMain_S.Nodes("K7").Tag = 7
            Call FillTree(7)
        End If
    End With
    If Me.tvwMain_S.Nodes.Count <> 0 Then
        Me.tvwMain_S.Nodes(1).Expanded = True
        Me.tvwMain_S.Nodes(1).Selected = True
        Call Me.tvwMain_S.Nodes(1).EnsureVisible
    End If
    
End Sub

Private Sub InitSum()
    '��ʼ�����ܱ����ʽ
    With mshSum
        ClearGrid mshSum, 9

'        .MergeCells = flexMergeRestrictRows
'        .MergeCol(col����) = True
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col����) = "�շ�ϸĿ"
        .TextMatrix(0, col���) = "���"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col�ۼ۵�λ) = "��λ"
        .TextMatrix(0, col������Ŀ) = "������Ŀ"
        .TextMatrix(0, col�۸�) = "�۸�"
        .TextMatrix(0, col����) = "�����շ���"
        .TextMatrix(0, col�Ӱ�) = "�Ӱ�Ӽ���"
        
        .ColWidth(col����) = 1000
        .ColWidth(col����) = 2500
        .ColWidth(col���) = 1600
        .ColWidth(col����) = 1500
        .ColWidth(col�ۼ۵�λ) = 600
        .ColWidth(col������Ŀ) = 900
        .ColWidth(col�۸�) = 1100
        .ColWidth(col����) = 800
        .ColWidth(col�Ӱ�) = 800
        
        .ColAlignment(col����) = 1
        .ColAlignment(col����) = 1
        .ColAlignment(col���) = 1
        .ColAlignment(col����) = 1
        .ColAlignment(col�ۼ۵�λ) = 1
        .ColAlignment(col������Ŀ) = 1
        .ColAlignment(col�۸�) = 7
        .ColAlignment(col����) = 7
        .ColAlignment(col�Ӱ�) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mrs��Ŀ.Sort = ""
    If mrs��Ŀ.State = 1 Then mrs��Ŀ.Close
    
    mstrKey = ""
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewShowDynamic״̬", mnuViewShowDynamic.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '�ұ�
    'tvwMain_S��λ��
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV��λ��
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
        
        
    mshSum.Left = picV.Left + picV.Width
    mshSum.Width = ScaleWidth - mshSum.Left
    mshSum.Top = sngTop
    
    If picHBar_S.Top > Me.ScaleHeight - 2000 Then picHBar_S.Top = Me.ScaleHeight - 2000
    picHBar_S.Left = mshSum.Left
    picHBar_S.Width = mshSum.Width
    If msh������Ŀ.Visible = False Then
        mshSum.Height = IIf(sngBottom - mshSum.Top > 0, sngBottom - mshSum.Top, 0)
    Else
        mshSum.Height = picHBar_S.Top - mshSum.Top '  IIf(sngBottom - mshSum.Top > 0, sngBottom - mshSum.Top, 0)
    End If
    With msh������Ŀ
        If .Visible = True Then
            .Left = mshSum.Left
            .Top = picHBar_S.Top + picHBar_S.Height
            .Height = IIf(sngBottom - .Top > 0, sngBottom - .Top, 0)
            .Width = Me.ScaleWidth - .Left
        End If
    End With
    Refresh
End Sub

Private Sub mnuViewFind_Click()
    With frmPriceFind
        .Left = Me.Left + Me.Width - .Width
        .Top = Me.Top + Me.Height - .Height
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    FillSum
End Sub

Private Sub mnuViewShowDynamic_Click()
    mnuViewShowDynamic.Checked = Not mnuViewShowDynamic.Checked
    
    mstrKey = ""
    Call FillSum
End Sub

 

Private Sub mshSum_GotFocus()
    Call MenuSet
    mshSum.BackColorSel = &H8000000D
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
    mshSum.BackColorSel = &H8000000F
End Sub

Private Sub mshSum_RowColChange()
    With mshSum
        If .Row <> mlngPreRow Then
            mlngPreRow = .Row
            
            Load������Ŀ .RowData(.Row)
        End If
    End With
End Sub

Private Sub msh������Ŀ_GotFocus()
    msh������Ŀ.BackColorSel = &H8000000D
End Sub

Private Sub msh������Ŀ_LostFocus()
    msh������Ŀ.BackColorSel = &H8000000F
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim intType As Integer
    mlngPreRow = -1
    Select Case Val(Node.Tag)
    Case 1, 2, 3        '����ҩ,�г�ҩ,�в�ҩ
        intType = 2     'intType-1- �շ���Ŀ,2-ҩƷ��Ŀ,3-������Ŀ
    Case 7              '��������
        intType = 3     'intType-1- �շ���Ŀ,2-ҩƷ��Ŀ,3-������Ŀ
    Case Else           '�շ���Ŀ
        intType = 1
    End Select
    mstrVbFormat = GetFmtString(intType, False)
    mstrOraFormat = GetFmtString(intType, True)
    FillSum
    Call Set������Ŀ
    Call mshSum_RowColChange
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tabMain_Click()
    mstrKey = ""
    Call FillSum
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
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
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hWnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum
    objPrint.Title.Text = nod.Text & "����Ŀ��Ŀ��"
    objRow.Add "ҽԺ���ƣ�" & gstr��λ����
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & gstrUserName
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
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

Private Function FillTree(lngKind As Long) As Boolean
    '����:װ���շ������շ�ϸĿ�����з��ൽtvwMain_S
    '�����������ڵ�����������KEYֵ��һ���ַ������ڶ�λ��������
    Dim objNode As Node
    
    Select Case lngKind
    Case 0
        gstrSQL = "Select id, �ϼ�id, ����, ���� " & _
                "  From �շѷ���Ŀ¼" & _
                " Start With �ϼ�ID Is Null" & _
                " Connect By Prior id = �ϼ�ID"
    Case 1, 2, 3, 7
        gstrSQL = "Select id, �ϼ�id, ����, ���� " & _
                "  From ���Ʒ���Ŀ¼" & _
                " Where ���� = " & lngKind & _
                " Start With �ϼ�ID Is Null" & _
                " Connect By Prior id = �ϼ�ID"
    End Select
    Call OpenRecordset(rsTemp, Me.Caption)
    With tvwMain_S.Nodes
        Do Until rsTemp.EOF
            If IsNull(rsTemp("�ϼ�id")) Then
                Set objNode = .Add("K" & lngKind, tvwChild, "K" & lngKind & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
            Else
                Set objNode = .Add("K" & lngKind & rsTemp("�ϼ�id"), tvwChild, "K" & lngKind & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "P", "P")
            End If
            objNode.Tag = lngKind
            objNode.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Set������Ŀ()
    '--------------------------------------------------------------------------------------------------
    '����:���ô�����Ŀ���������
    '����:���˺�
    '����:2007/09/27
    '--------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Dim blnOldHide As Boolean       '�ϴ��Ƿ�Ӱ�ص�
    
    If Me.tvwMain_S.SelectedItem Is Nothing Then
        blnVisible = False
    Else
        blnVisible = Val(Me.tvwMain_S.SelectedItem.Tag) = 0
    End If
    blnOldHide = msh������Ŀ.Visible
    
    With msh������Ŀ
        .Visible = blnVisible
        picHBar_S.Visible = blnVisible
    End With
    If blnOldHide = False And blnVisible Then
        If picHBar_S.Top < Me.ScaleHeight - picHBar_S.Top - IIf(stbThis.Visible = False, 0, stbThis.Height) Then
            picHBar_S.Top = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000
        End If
    End If
    If picHBar_S.Top > Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000 Then
         picHBar_S.Top = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000
    End If
    Call Form_Resize
End Sub
Public Sub FillSum()
    '����:װ�����ͳ������
    Dim nod As Node
    Dim str���ʷ��� As String

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum
        Call MenuSet
        Exit Sub
    End If
    If mstrKey = tvwMain_S.SelectedItem.Key Then Exit Sub
    mstrKey = tvwMain_S.SelectedItem.Key
    Set nod = tvwMain_S.SelectedItem
    
    '���ݲ�ͬ�Ľڵ㣬������ͬ����ʾ
    Select Case Mid(nod.Key, 2, 1)
    Case "0"
        mshSum.TextMatrix(0, col����) = "����"
        mshSum.ColWidth(col����) = 0
        mshSum.ColWidth(col����) = 1100
        mshSum.ColWidth(col�Ӱ�) = 1000
        mshSum.MergeCol(col����) = True
        mshSum.MergeCol(col����) = True
        
        If nod.Image = "R" Then
            gstrSQL = "Select id,����,����,�Ӱ�Ӽ�,�Ƿ���,���,����,���㵥λ" & _
                    " From �շ���ĿĿ¼" & _
                    " Where ��� not in ('4','5','6','7') And " & IIf(mnuViewShowDynamic.Checked, "", " �Ƿ���=0 And ") & _
                    "       (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "Select id,����,����,�Ӱ�Ӽ�,�Ƿ���,���,����,���㵥λ" & _
                    " From �շ���ĿĿ¼" & _
                    " Where ��� not in ('4','5','6','7') And " & IIf(mnuViewShowDynamic.Checked, "", " �Ƿ���=0 And ") & _
                    "       (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd')) And" & _
                    "       ����ID in (" & _
                    "           Select Id From �շѷ���Ŀ¼ start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID)"
        End If
        gstrSQL = "select A.�շ�ϸĿID,C.����,C.���� as �շ�ϸĿ,C.���,C.����,C.���㵥λ as �ۼ۵�λ,B.���� as ������Ŀ,A.ԭ��,A.�ּ�,A.�����շ���,decode(C.�Ӱ�Ӽ�,1,A.�Ӱ�Ӽ���,0) as �Ӱ�Ӽ���,C.�Ƿ��� " & _
                   " from �շѼ�Ŀ A,������Ŀ B, " & _
                   "   (" & gstrSQL & ") C " & _
                   " Where A.�շ�ϸĿID = C.ID And A.������ĿID = B.ID " & _
                   "       and A.ִ������<=sysdate and (A.��ֹ����>=sysdate or a.��ֹ���� is null) " & _
                   " order by C.����"
    
    Case "1", "2", "3", "7"
        Select Case Mid(nod.Key, 2, 1)
        Case "1", "2"
            mshSum.TextMatrix(0, col����) = "����"
            mshSum.ColWidth(col����) = 1500
            mshSum.ColWidth(col���) = 1600
        Case "3"
            mshSum.TextMatrix(0, col����) = "����"
            mshSum.ColWidth(col����) = 1000
            mshSum.ColWidth(col���) = 0
        Case "7"
            mshSum.TextMatrix(0, col����) = "����"
            mshSum.ColWidth(col����) = 1000
            mshSum.ColWidth(col���) = 1600
        End Select
        mshSum.ColWidth(col�Ӱ�) = 0
        mshSum.ColWidth(col����) = 0
        mshSum.MergeCol(col����) = False
        mshSum.MergeCol(col����) = False
        
        If nod.Image = "R" Then
            If mblnTradeName = False Then
                gstrSQL = "Select id,����,����,�Ӱ�Ӽ�,�Ƿ���,���,����,���㵥λ" & _
                        " From �շ���ĿĿ¼" & _
                        " Where ��� ='" & Switch(nod.Key = "K1", 5, nod.Key = "K2", 6, nod.Key = "K3", 7, nod.Key = "K7", 4) & "'" & _
                        "       And " & IIf(mnuViewShowDynamic.Checked, "", " �Ƿ���=0 And ") & _
                        "       (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                gstrSQL = "Select id, ����, nvl(N.����,I.����) As ����, �Ӱ�Ӽ�, �Ƿ���, ���, ����, ���㵥λ" & _
                        " From �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
                        " Where ��� ='" & Switch(nod.Key = "K1", 5, nod.Key = "K2", 6, nod.Key = "K3", 7, nod.Key = "K7", 4) & "'" & _
                        "       And " & IIf(mnuViewShowDynamic.Checked, "", " �Ƿ���=0 And ") & _
                        "       (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        "       And I.Id=N.�շ�ϸĿid(+) And N.����(+)=3 And N.����(+)=1"
            End If
        Else
            If mblnTradeName = False Then
                If Mid(nod.Key, 2, 1) = 7 Then
                    gstrSQL = "Select I.id, I.����, I.����, I.�Ӱ�Ӽ�, I.�Ƿ���, I.���, I.����, I.���㵥λ" & _
                            "  From �շ���ĿĿ¼ I,�������� T,������ĿĿ¼ Z" & _
                            " Where I.��� ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.�Ƿ���=0 And ") & _
                            "       (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=T.����id And T.����id=Z.Id" & _
                            "       And Z.����ID in (" & _
                            "           Select Id From ���Ʒ���Ŀ¼ start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID)"
                Else
                    gstrSQL = "Select I.id, I.����, I.����, I.�Ӱ�Ӽ�, I.�Ƿ���, I.���, I.����, I.���㵥λ" & _
                            "  From �շ���ĿĿ¼ I,ҩƷ��� T,������ĿĿ¼ Z" & _
                            " Where I.��� ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.�Ƿ���=0 And ") & _
                            "       (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=T.ҩƷid And T.ҩ��id=Z.Id" & _
                            "       And Z.����ID in (" & _
                            "           Select Id From ���Ʒ���Ŀ¼ start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID)"
                End If
            Else
                If Mid(nod.Key, 2, 1) = 7 Then
                    gstrSQL = "Select I.id, I.����, nvl(N.����,I.����) As ����, I.�Ӱ�Ӽ�, I.�Ƿ���, I.���, I.����, I.���㵥λ" & _
                            "  From �շ���ĿĿ¼ I,�շ���Ŀ���� N,�������� T,������ĿĿ¼ Z" & _
                            " Where I.��� ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.�Ƿ���=0 And ") & _
                            "       (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=N.�շ�ϸĿid(+) And N.����(+)=3 And N.����(+)=1" & _
                            "       And I.Id=T.����id And T.����id=Z.Id" & _
                            "       And Z.����ID in (" & _
                            "           Select Id From ���Ʒ���Ŀ¼ start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID)"
                Else
                    gstrSQL = "Select I.id, I.����, nvl(N.����,I.����) As ����, I.�Ӱ�Ӽ�, I.�Ƿ���, I.���, I.����, I.���㵥λ" & _
                            "  From �շ���ĿĿ¼ I,�շ���Ŀ���� N,ҩƷ��� T,������ĿĿ¼ Z" & _
                            " Where I.��� ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.�Ƿ���=0 And ") & _
                            "       (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=N.�շ�ϸĿid(+) And N.����(+)=3 And N.����(+)=1" & _
                            "       And I.Id=T.ҩƷid And T.ҩ��id=Z.Id" & _
                            "       And Z.����ID in (" & _
                            "           Select Id From ���Ʒ���Ŀ¼ start with ID=" & Mid(nod.Key, 3) & " connect by prior id=�ϼ�ID)"
                End If
            End If
        End If
        gstrSQL = "select A.�շ�ϸĿID,C.����,C.���� as �շ�ϸĿ,C.���,C.����,C.���㵥λ as �ۼ۵�λ,B.���� as ������Ŀ,A.ԭ��,A.�ּ�,A.�����շ���,decode(C.�Ӱ�Ӽ�,1,A.�Ӱ�Ӽ���,0) as �Ӱ�Ӽ���,C.�Ƿ��� " & _
                   " from �շѼ�Ŀ A,������Ŀ B, " & _
                   "   (" & gstrSQL & ") C " & _
                   " Where A.�շ�ϸĿID = C.ID And A.������ĿID = B.ID " & _
                   "       and A.ִ������<=sysdate and (A.��ֹ����>=sysdate or a.��ֹ���� is null) " & _
                   " order by C.����"
    End Select
    
    MousePointer = 11
    If mrs��Ŀ.State = 1 Then mrs��Ŀ.Close
    Call OpenRecordset(mrs��Ŀ, Me.Caption)
    
    Call ReList
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub ReList()
    Dim lngRow As Long
    Dim lngID  As Long
    Dim lngCount As Long
    
    
    MousePointer = 11
    mshSum.Redraw = False
    ClearGrid mshSum
    If mrs��Ŀ.RecordCount <> 0 Then
        mshSum.Rows = mrs��Ŀ.RecordCount + 1
    End If
    lngRow = 1
    With mshSum
        Do Until mrs��Ŀ.EOF
            If mrs��Ŀ("�շ�ϸĿID") <> lngID Then
                lngID = mrs��Ŀ("�շ�ϸĿID")
                lngCount = lngCount + 1
            End If
            .RowData(lngRow) = lngID
            .TextMatrix(lngRow, col����) = mrs��Ŀ("����")
            .TextMatrix(lngRow, col����) = mrs��Ŀ("�շ�ϸĿ")
            .TextMatrix(lngRow, col���) = IIf(IsNull(mrs��Ŀ("���")), "", mrs��Ŀ("���"))
            .TextMatrix(lngRow, col����) = IIf(IsNull(mrs��Ŀ("����")), "", mrs��Ŀ("����"))
            .TextMatrix(lngRow, col�ۼ۵�λ) = IIf(IsNull(mrs��Ŀ("�ۼ۵�λ")), "", mrs��Ŀ("�ۼ۵�λ"))
            .TextMatrix(lngRow, col������Ŀ) = mrs��Ŀ("������Ŀ")
            If mrs��Ŀ("�Ƿ���") = 1 Then
                .TextMatrix(lngRow, col�۸�) = Format(mrs��Ŀ("ԭ��"), mstrVbFormat) & "��" & Format(mrs��Ŀ("�ּ�"), mstrVbFormat)
            Else
                .TextMatrix(lngRow, col�۸�) = Format(mrs��Ŀ("�ּ�"), mstrVbFormat)
            End If
            .TextMatrix(lngRow, col����) = Format(mrs��Ŀ("�����շ���"), "0.00;-0.00; ; ")
            .TextMatrix(lngRow, col�Ӱ�) = Format(mrs��Ŀ("�Ӱ�Ӽ���"), "0.00;-0.00; ; ")
            lngRow = lngRow + 1
            mrs��Ŀ.MoveNext
        Loop
    End With
    mshSum.Redraw = True
    stbThis.Panels(2).Text = "�����շ���Ŀ" & lngCount & "��"
    MousePointer = 0

End Sub

Private Sub ClearGrid(objGrid As MSHFlexGrid, Optional lngCols As Long = 0)
'���ܣ�������,����ɲ��ֳ�ʼ��
    Dim i As Long
    
    With objGrid
        If lngCols > 0 Then
            '������������������Ǿͳ�ʼ����
            .Cols = lngCols
            .AllowBigSelection = True
            .FillStyle = flexFillRepeat
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = 0
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
            .Row = 1
        End If
        
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'����:��ʾ�˵��͹�������״̬(��ӡ)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum.Rows = 2 And mshSum.TextMatrix(1, col����) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub






Private Sub picHBar_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    msngOldY = y
End Sub

Private Sub picHBar_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picHBar_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    
    Call Form_Resize
    
    
End Sub

Private Sub picHBar_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        msngOldY = 0
End Sub
Private Sub Load������Ŀ(ByVal lng�շ�ϸĿID As Long)
    '---------------------------------------------------------------------------------------
    '����:����������Ŀ
    '����:
    '����:���˺�
    '����:2007/09/28
    '---------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim j As Integer, i As Integer, intCol As Integer, lngRow As Long
    
    If msh������Ŀ.Visible = False Then Exit Sub
    
    
    gstrSQL = "" & _
        "   Select a.����ID,a.����ID,a.���д���,a.��������,b.����,b.���� ��Ŀ����,c.���� ,c.���� ���, " & vbCrLf & _
        "           Nvl(B.����ʱ��,to_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��," & vbCrLf & _
        "           decode(nvl(b.�Ƿ���,0),1,ltrim(rtrim(to_char(sum(d.ԭ��),'" & mstrOraFormat & "')))||'��'||ltrim(rtrim(to_char(sum(d.�ּ�),'" & mstrOraFormat & "'))),ltrim(rtrim(to_char(sum(d.�ּ�),'" & mstrOraFormat & "'))))  AS  �۸� " & vbCrLf & _
        "   From �շѴ�����Ŀ a,�շ���ĿĿ¼ b,�շ���Ŀ��� c ,�շѼ�Ŀ d " & vbCrLf & _
        "   Where c.����=b.��� and  a.����ID=b.id  and b.id=d.�շ�ϸĿid  and ����ID=[1] " & vbCrLf & _
        "           AND NVL (D.��ֹ����, TO_DATE ('3000-01-01', 'YYYY-MM-DD')) = TO_DATE ('3000-01-01', 'YYYY-MM-DD') " & _
        "   GROUP BY a.ROWID,a.����ID,b.�Ƿ���,a.����ID,a.���д���,a.��������,b.����,b.����,b.����ʱ�� ,c.���� ,c.���� " & _
        " ORDER BY a.ROWID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�շ�ϸĿID)
    With msh������Ŀ
        .Redraw = flexRDNone
        If rsTemp.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
            .Redraw = flexRDBuffered
            Exit Sub
        End If
        .Rows = rsTemp.RecordCount + 1
        Dim dbl���� As Double, intTemp As Integer
        
        i = 1
        dbl���� = 0
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("�շ����")) = "(" & NVL(rsTemp!����) & ")" & NVL(rsTemp!���)
            .TextMatrix(i, .ColIndex("�շ���Ŀ")) = "[" & NVL(rsTemp!��Ŀ����) & "]" & NVL(rsTemp!����)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!��������)
            intTemp = Val(NVL(rsTemp!���д���))
            If intTemp = 0 Then
                .TextMatrix(i, .ColIndex("�̶�")) = "0-���̶�"
            ElseIf intTemp = 2 Then
                .TextMatrix(i, .ColIndex("�̶�")) = "2-����������"
            Else
                .TextMatrix(i, .ColIndex("�̶�")) = "1-�̶�"
            End If
            
            If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                lngRow = .Row: intCol = .Col
                .Row = i
                For j = 0 To .Cols - 1
                    .Col = j
                    .CellForeColor = &HFF&
                Next
                .Row = lngRow: .Col = intCol
                .TextMatrix(i, .ColIndex("״̬")) = "ͣ��"
            Else
                .TextMatrix(i, .ColIndex("״̬")) = ""
            End If
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp("�۸�"))
            If IsNumeric(.TextMatrix(i, .ColIndex("����"))) Then
                dbl���� = dbl���� + Val(.TextMatrix(i, .ColIndex("����")))
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '����ϼ���
        .Rows = .Rows + 1
        .TextMatrix(i, .ColIndex("�շ����")) = ""
        .TextMatrix(i, .ColIndex("�շ���Ŀ")) = "�ϼ�"
        .TextMatrix(i, .ColIndex("����")) = ""
        .TextMatrix(i, .ColIndex("�̶�")) = ""
        .TextMatrix(i, .ColIndex("״̬")) = ""
        .TextMatrix(i, .ColIndex("����")) = Format(dbl����, mstrVbFormat)
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function NVL(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '����:ȡĳ�ֶε�ֵ
    '����:rsObj          �������ֶ�
    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
    End If
End Function

Private Sub initС��λ��()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    '    ���    Number(1)   1-ҩƷ,2-����
    '    ����    Number(1)   1-�ɱ��ۣ�2-���ۼ�,3-����
    '    ��λ    Number(1)   1,2,3,4��ҩƷ�ֱ�Ϊ�ۼۡ����סԺ��ҩ�ⵥλ�����ķֱ�Ϊɢװ����װ��λ��
    '    ����    Number(1)   ȡֵΪ2-7��
    
    
    m�ۼ۾���.ҩƷ��ĿС�� = 7
    m�ۼ۾���.������ĿС�� = 7
    m�ۼ۾���.�շ���ĿС�� = 3
    strSQL = "Select * from ҩƷ���ľ��� where ��� in (1,2) and ����=2 and  ��λ=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������ϵ�С��λ������", 2)
    Do While Not rsTemp.EOF
        Select Case Val(NVL(rsTemp!���))
        Case 1
            m�ۼ۾���.ҩƷ��ĿС�� = Val(NVL(rsTemp!����, "7"))
        Case 2
            m�ۼ۾���.������ĿС�� = Val(NVL(rsTemp!����, "7"))
        End Select
        rsTemp.MoveNext
    Loop
End Sub
Private Function GetFmtString(ByVal intType As Integer, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '����:����ָ����С����ʽ��
    '���:intType-1- �շ���Ŀ,2-ҩƷ��Ŀ,3-������Ŀ
    '     blnOracle-������oracle�ĸ�ʽ������Vb�ĸ�ʽ��
    '����:
    '����:����ָ���ĸ�ʽ��
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim intλ�� As Integer
    Select Case intType
    Case 2  '-ҩƷ��Ŀ
         intλ�� = m�ۼ۾���.ҩƷ��ĿС��
    Case 3  '������Ŀ
         intλ�� = m�ۼ۾���.������ĿС��
    Case Else       '�շ���Ŀ
         intλ�� = m�ۼ۾���.�շ���ĿС��
    End Select
    If blnOracle Then
       GetFmtString = "99999999999990." & String(intλ��, "9") & ""
    Else
       GetFmtString = "#0." & String(intλ��, "0") & ";-#0." & String(intλ��, "0") & ";0; "
    End If
End Function
 
