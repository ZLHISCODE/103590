VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmӦ�����ѯ 
   BackColor       =   &H8000000A&
   Caption         =   "ҩƷӦ�����ѯ"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frmӦ�����ѯ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   635
      SimpleText      =   $"frmӦ�����ѯ.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmӦ�����ѯ.frx":0351
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9340
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
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   345
      Left            =   4710
      TabIndex        =   8
      Top             =   2880
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   609
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������ϸ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ѹ��嵥"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "δ���嵥"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   1005
      Left            =   3330
      TabIndex        =   7
      Top             =   1290
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1773
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2550
      Width           =   3000
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2190
      Top             =   2820
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
            Picture         =   "frmӦ�����ѯ.frx":0BE5
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":0D3F
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":0E99
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":11B3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":13CF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":15EB
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":1805
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":1A1F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":1C3B
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":1E57
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":2073
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":228F
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":24A9
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":26C3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�����ѯ.frx":28DF
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8160
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
               Key             =   "Open"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Find"
               Description     =   "��λ����"
               Object.ToolTipText     =   "��λ"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1215
      Left            =   2910
      TabIndex        =   9
      Top             =   3240
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ϸ��Ϣ"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3255
      TabIndex        =   10
      Top             =   2610
      Width           =   3015
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������Ϣ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3450
      TabIndex        =   4
      Top             =   930
      Width           =   1950
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
      Begin VB.Menu mnuViewUnit 
         Caption         =   "�ۼ۵�λ(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "���ﵥλ(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "סԺ��λ(&Z)"
         Index           =   2
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "ҩ�ⵥλ(&K)"
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "��������(&J)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "���ݶ�λ(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
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
Attribute VB_Name = "frmӦ�����ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnLoad As Boolean
Dim mdatBegin As Date, mdatEnd As Date            '��ѯ��ʱ�䷶Χ
Dim mstrData As String
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Dim mlngID As Long          'ǰһ��ҩƷ��Ӧ�̵�ID

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillTree
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim i As Double
    mblnLoad = True
    RestoreWinState Me, App.ProductName
    
    i = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewUnit", "0"))
    If i < 0 Or i > 3 Then
        i = 0
    Else
        i = Int(i)
    End If
    mnuViewUnit(i).Checked = True
    
    If glngSys \ 100 = 8 Then
        'ҩ��ϵͳ
        mnuViewUnit(1).Visible = False
        mnuViewUnit(2).Visible = False
        mnuViewUnit(3).Caption = "�ɹ���λ(&K)"
    End If
    '�õ���ѯ��ʱ�䷶Χ
    mdatEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    mdatBegin = DateAdd("m", -1, mdatEnd) + 1
    mstrData = "0000"
    Call InitSum
End Sub

Private Sub InitSum()
'��ʼ�����ܱ����ʽ
    With mshSum
        ClearGrid mshSum, 5
        .TextMatrix(0, 0) = "ҩƷ��Ӧ��"
        .TextMatrix(0, 1) = "�ڳ�Ӧ��"
        .TextMatrix(0, 2) = "�����޹�"
        .TextMatrix(0, 3) = "����֧��"
        .TextMatrix(0, 4) = "��ĩӦ��"
        
        .colWidth(0) = 2000
        .colWidth(1) = 1500
        .colWidth(2) = 1500
        .colWidth(3) = 1500
        .colWidth(4) = 1500
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    Dim i As Integer
    
    For i = 0 To 3
        If mnuViewUnit(i).Checked = True Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewUnit", i
        End If
    Next
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
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
    '���
    lblSum.Top = sngTop
    lblSum.Left = picV.Left + picV.Width
    If ScaleWidth - lblSum.Left > 0 Then lblSum.Width = ScaleWidth - lblSum.Left
    
    mshSum.Left = lblSum.Left
    picH.Left = lblSum.Left
    lblDetail.Left = lblSum.Left
    tabMain.Left = lblDetail.Left + lblDetail.Width + 60
    mshDetail.Left = lblSum.Left
    
    mshSum.Width = lblSum.Width
    picH.Width = lblSum.Width
    tabMain.Width = ScaleWidth - tabMain.Left
    mshDetail.Width = lblSum.Width
    
    'mshSum��λ��
    mshSum.Top = lblSum.Top + lblSum.Height
    'picH��λ��
    picH.Top = mshSum.Top + mshSum.Height
    'tabMain��λ��
    lblDetail.Top = picH.Top + picH.Height + 15
    tabMain.Top = picH.Top + picH.Height
    'mshDetail��λ��
    mshDetail.Top = tabMain.Top + tabMain.Height
    mshDetail.Height = IIf(sngBottom - mshDetail.Top > 0, sngBottom - mshDetail.Top, 0)
    
    Refresh
End Sub

Private Sub mnuViewFind_Click()
'��ҩƷ��Ӧ���뵥�ݺŶ�λ
    Dim str���ݺ� As String, str��Ӧ��ID As String
    Dim rsTemp As New ADODB.Recordset
    Dim nod As MSComctlLib.node, lngRow As Long, lngCol As Long
    
    If frmӦ���λ.Get��λ����(str���ݺ�, str��Ӧ��ID) = False Then
        Exit Sub
    End If
    
    If str���ݺ� <> "" Then
        '���ݵ��ݺ��ҵ���Ӧ��
        gstrSQL = "select ��ҩ��λID from ҩƷ�շ���¼ where NO='" & str���ݺ� & "' and ����=1 and ��ҩ��λID is not null"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = True Then
            MsgBox "���ݺ�Ϊ " & str���ݺ� & " �⹺��ⵥû���ҵ���", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str��Ӧ��ID = rsTemp("��ҩ��λID")
        rsTemp.Close
    End If
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes("C" & str��Ӧ��ID)
    If Err <> 0 Then
        MsgBox "û�з���ָ����Ӧ�̣������Ѿ���ͣ�á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    nod.Selected = True
    nod.EnsureVisible
    Call FillSum
    
    If str���ݺ� <> "" Then
        '�ҵ�����������
        If tabMain.SelectedItem.Index = 1 Then
            lngCol = 1
        Else
            lngCol = 0
        End If
        
        With mshDetail
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, lngCol) = str���ݺ� Then
                    .TopRow = lngRow
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub mnuViewOpen_Click()
    If frmTimeSet.GetTimeScope(mdatBegin, mdatEnd, mstrData, Me) = True Then
        mstrKey = ""
        Call FillSum
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FillTree
End Sub

Private Sub mnuViewUnit_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewUnit(i).Checked = False
    Next
    mnuViewUnit(Index).Checked = True
    
    Call FillDetail
End Sub

Private Sub mshSum_EnterCell()
    If mlngID = mshSum.RowData(mshSum.Row) Then Exit Sub
    mlngID = mshSum.RowData(mshSum.Row)
    Call FillDetail
End Sub

Private Sub mshSum_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal node As MSComctlLib.node)
    FillSum
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
        If sngTemp > 500 And ScaleWidth - (sngTemp + picV.Width) > 600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub picH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picH.Top + y - msngStartY
        If sngTemp > mshSum.Top + 600 And ScaleHeight - (sngTemp + picH.Height) > 1600 Then
            picH.Top = sngTemp
            mshSum.Height = picH.Top - mshSum.Top
            Form_Resize
        End If
    End If
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub


Private Sub tabMain_Click()
    Call FillDetail
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            mnuViewOpen_Click
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnufileexit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreView_Click
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
   Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If mshSum Is ActiveControl Then
        Set objPrint.Body = mshSum
        objPrint.Title.Text = "Ӧ����ҩ�������Ϣ"
        objRow.Add " "
        objRow.Add "��ѯʱ�䣺" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.�û�����
        objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = mshDetail
        objPrint.Title.Text = tabMain.SelectedItem.Caption
        objRow.Add "ҩƷ��Ӧ�̣�" & lblDetail.Caption
        objRow.Add "��ѯʱ�䣺" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.�û�����
        objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    End If
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

Private Function FillTree() As Boolean
'����:װ��ҩƷ��Ӧ��
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strkey As String
    
    mstrKey = ""     'ȫ��ˢ��ʱ���൱���û�û����κνڵ�
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strkey = tvwMain_S.SelectedItem.Key
    End If
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select id,�ϼ�id,����,����,ĩ�� from ҩƷ��Ӧ��  where ����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null " & _
         " start with �ϼ�ID is null  connect by prior id=�ϼ�ID "
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "��ҩƷ��Ӧ�̡�����Ϣ��ȫ���޷����в�ѯ��", vbExclamation, gstrSysName
        FillTree = False
        Exit Function
    End If
    
    
    With tvwMain_S.Nodes
        .Clear
        .Add , , "Root", "����ҩƷ��Ӧ��", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        Do Until rsTemp.EOF
            '�ó���ȷ��ͼ��
            strTemp = IIf(rsTemp("ĩ��") = 1, "Item", "Class")
            '��ӽڵ�
            If IsNull(rsTemp("�ϼ�id")) Then
                .Add "Root", tvwChild, "C" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), strTemp, strTemp
            Else
                .Add "C" & rsTemp("�ϼ�id"), tvwChild, "C" & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), strTemp, strTemp
            End If
            tvwMain_S.Nodes("C" & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    Dim nod As node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strkey)
    If Err <> 0 Then
        Err.Clear
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
    Else
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
    End If
    Call FillSum
    FillTree = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillSum()
'����:װ�����ͳ������
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 4) As Double
    Dim lngRow As Long
    Dim blnSum As Boolean        '�ϼƵ���ʾ
    
    
    stbThis.Panels(2).Text = "ʱ�䷶Χ��" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")

    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    If mstrKey = tvwMain_S.SelectedItem.Key Then Exit Sub
    mstrKey = tvwMain_S.SelectedItem.Key
    '��ʼ��ѯ
    
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd, "yyyyMMdd")
    MousePointer = 11
    '���ȵõ��Ӳ�ѯ��SQL���
    If tvwMain_S.SelectedItem.Image = "Item" Then
        gstrSQL = " and A.��λID=" & Mid(mstrKey, 2)
    ElseIf tvwMain_S.SelectedItem.Image = "Root" Then
        gstrSQL = " and A.��λID in (select ID from ҩƷ��Ӧ�� start with �ϼ�ID is null connect by prior id=�ϼ�ID)"
    Else
        gstrSQL = " and A.��λID in (select ID from ҩƷ��Ӧ�� start with �ϼ�ID =" & Mid(mstrKey, 2) & " connect by prior id=�ϼ�ID)"
    End If
    If Mid(mstrData, 1, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.�ڳ�Ӧ��<>0 "
    End If
    If Mid(mstrData, 2, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.�����޹�<>0 "
    End If
    If Mid(mstrData, 3, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.����֧��<>0 "
    End If
    If Mid(mstrData, 4, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.��ĩӦ��<>0 "
    End If
    
    '�ٵõ�������SQL���
    gstrSQL = "select B.����,B.ID,A.�ڳ�Ӧ��,A.�����޹�,A.����֧��,A.��ĩӦ�� from " & _
            "(select ��λID,sum(���-�ڳ�Ӧ��+�ڳ�����) as �ڳ�Ӧ��,sum(�ڳ�Ӧ��-��ĩӦ��) as �����޹� " & _
            "            ,sum(�ڳ�����-��ĩ����) as ����֧��,sum(���-��ĩӦ��+��ĩ����) as ��ĩӦ�� " & _
            "from( " & _
            "select ��λID,��� as �ڳ�����, " & _
            "    decode(sign(to_char(�������,'yyyymmdd')-'" & strEnd & "'),1,���,0) as ��ĩ����, " & _
            "    0 as �ڳ�Ӧ��,0 as ��ĩӦ��,0 as ��� from ҩƷ�����¼ " & _
            "    where �������>=to_date('" & strBegin & "','yyyyMMdd') " & _
            "Union All " & _
            "select A.��ҩ��λID ��λID,0 as �ڳ�����,0 as ��ĩ����, " & _
            "    A.��Ʊ��� as �ڳ�Ӧ��,decode(sign(to_char(B.�������,'yyyymmdd')-'" & strEnd & "'),1,A.��Ʊ���,0) as ��ĩӦ��,0 as ��� from ҩƷӦ����¼ A,ҩƷ�շ���¼ B " & _
            "    where A.�շ�ID=B.ID and B.�������>=to_date('" & strBegin & "','yyyyMMdd') " & _
            "Union All " & _
            "select ��ҩ��ID ��λID,0 as �ڳ�����,0 as ��ĩ����,0 as �ڳ�Ӧ��,0 as ��ĩӦ��,��� as ��� from ҩƷӦ����� " & _
            "    where ����=1) " & _
            "group by ��λID)A,ҩƷ��Ӧ�� B " & _
            "where A.��λID=B.ID  " & gstrSQL
    Call OpenRecordset(rsTemp, Me.Caption)
    
    mshSum.Redraw = False
    If rsTemp.RecordCount = 0 Then
        ClearGrid mshSum
    Else
        If rsTemp.RecordCount = 1 Then
            'ֻ��һ�У��Ͳ���ʾ�ϼ���
            mshSum.Rows = 2
            blnSum = False
        Else
            mshSum.Rows = rsTemp.RecordCount + 2
            blnSum = True
        End If
    End If
    lngRow = 1
    With mshSum
        Do Until rsTemp.EOF
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, 0) = rsTemp("����")
            .TextMatrix(lngRow, 1) = Format(rsTemp("�ڳ�Ӧ��"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("�����޹�"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 3) = Format(rsTemp("����֧��"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 4) = Format(rsTemp("��ĩӦ��"), "###########0.00;-###########0.00; ; ")
            If blnSum = True Then
                dblSum(1) = dblSum(1) + rsTemp("�ڳ�Ӧ��")
                dblSum(2) = dblSum(2) + rsTemp("�����޹�")
                dblSum(3) = dblSum(3) + rsTemp("����֧��")
                dblSum(4) = dblSum(4) + rsTemp("��ĩӦ��")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If blnSum = True Then
            .TextMatrix(lngRow, 0) = "  �ϼ�"
            .TextMatrix(lngRow, 1) = Format(dblSum(1), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(dblSum(2), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 3) = Format(dblSum(3), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 4) = Format(dblSum(4), "###########0.00;-###########0.00; ; ")
        End If
    End With
    mshSum.Redraw = True
    
    MousePointer = 0
    Call FillDetail
End Sub

Private Sub FillDetail()
'����:װ�������ϸ����
    
    MousePointer = 11
    If mshSum.RowData(mshSum.Row) <> 0 Then
        lblDetail.Caption = mshSum.TextMatrix(mshSum.Row, 0)
    Else
        lblDetail.Caption = "��ϸ��Ϣ"
    End If
    lblDetail.ToolTipText = lblDetail.Caption
    
    mshDetail.Redraw = False
    Select Case tabMain.SelectedItem.Index
        Case 2
            Call Fill�Ѹ��嵥
        Case 3
            Call Fillδ���嵥
        Case Else
            Call Fill��ϸ��
    End Select
    mshDetail.Redraw = True
    MousePointer = 0
    
    Call MenuSet
End Sub

Private Sub Fill��ϸ��()
'����:װ����ϸ������
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 2) As Double, dblBalance As Double
    Dim lngRow As Long, lngID As Long
    
    '��ʼ�����
    ClearGrid mshDetail, 11
    
    With mshDetail
        .MergeCells = flexMergeNever
        .TextMatrix(0, 0) = "����":     .colWidth(0) = 1100: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "���ݺ�":   .colWidth(1) = 1000: .ColAlignment(1) = 1
        .TextMatrix(0, 2) = "ժҪ":     .colWidth(2) = 1750: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "��λ":     .colWidth(3) = 600: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "����":     .colWidth(4) = 900: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "ʧЧ��":     .colWidth(5) = 1100: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "�ɹ�����": .colWidth(6) = 900: .ColAlignment(6) = 7
        .TextMatrix(0, 7) = "�ɹ���":   .colWidth(7) = 900: .ColAlignment(7) = 7
        .TextMatrix(0, 8) = "Ӧ�����": .colWidth(8) = 1000: .ColAlignment(8) = 7
        .TextMatrix(0, 9) = "�Ѹ����": .colWidth(9) = 1000: .ColAlignment(9) = 7
        .TextMatrix(0, 10) = "���":     .colWidth(10) = 1000: .ColAlignment(10) = 7
    End With
    '�õ���ѯ����
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '��ʼ��ѯ
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    '���ȵõ���ĩ���
    gstrSQL = "select sum(���) as ��� from( " & _
            "select ���  from ҩƷ�����¼ " & _
            "    where �������>=to_date('" & strEnd & "','yyyyMMdd')  and ��λID=" & lngID & _
            " Union All " & _
            "select -1 * A.��Ʊ��� as ��� from ҩƷӦ����¼ A,ҩƷ�շ���¼ B " & _
            "    where A.�շ�ID=B.ID and B.�������>=to_date('" & strEnd & "','yyyyMMdd')  and A.��ҩ��λID=" & lngID & _
            " Union All " & _
            "select ���  from ҩƷӦ����� " & _
            "    where ����=1 and ��ҩ��ID=" & lngID & ") "
    Call OpenRecordset(rsTemp, Me.Caption)
    
    dblBalance = IIf(IsNull(rsTemp("���")), 0, rsTemp("���"))
    rsTemp.Close
    
    
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",D.���ﵥλ as ��λ,B.ʵ������/D.�����װ as �ɹ�����,B.�ɱ���*D.�����װ as �ɹ���"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",D.סԺ��λ as ��λ,B.ʵ������/D.סԺ��װ as �ɹ�����,B.�ɱ���*D.סԺ��װ as �ɹ���"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",D.ҩ�ⵥλ as ��λ,B.ʵ������/D.ҩ���װ as �ɹ�����,B.�ɱ���*D.ҩ���װ as �ɹ���"
    Else
        gstrSQL = ",D.�ۼ۵�λ as ��λ,B.ʵ������ as �ɹ�����,B.�ɱ��� as �ɹ���"
    End If
    '�ٵõ���ϸ��
    gstrSQL = " select * from( " & _
              "  select to_char(�������,'yyyy-MM-dd') as ����,'��'||NO as NO,���, " & _
              "       decode(Ԥ����,1,'Ԥ����',decode(��¼״̬,2,'Ԥ����',ժҪ))||'('||���㷽ʽ||')' as ժҪ, " & _
              "       '' as ����,'' as Ч��,'' as ��λ,0 as �ɹ�����,0 as �ɹ���,0 as Ӧ�����,��� as �Ѹ���� " & _
              "       From ҩƷ�����¼ " & _
              "       where �������>=to_date('" & strBegin & "','yyyyMMdd') and �������<to_date('" & strEnd & "','yyyyMMdd') and ��λID=" & lngID & _
              "  Union All " & _
              "  select to_char(B.�������,'yyyy-MM-dd') as ����,B.NO,B.���,C.ͨ������||'('||D.���||')' as ժҪ, " & _
              "       ����,to_char(Ч��,'yyyy-MM-dd') as Ч��" & gstrSQL & ",A.��Ʊ��� as Ӧ�����,0 as �Ѹ���� " & _
              "       from ҩƷӦ����¼ A,ҩƷ�շ���¼ B,ҩƷĿ¼ D,ҩƷ��Ϣ C " & _
              "       where B.�������>=to_date('" & strBegin & "','yyyyMMdd') and B.�������<to_date('" & strEnd & "','yyyyMMdd') and A.��ҩ��λID=" & lngID & _
              "             and A.�շ�ID = B.ID And B.ҩƷID=D.ҩƷID and C.ҩ��ID=D.ҩ��ID ) " & _
              "  order by ����,no,���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    mshDetail.Rows = rsTemp.RecordCount + 3
    lngRow = 2
    With mshDetail
        .TextMatrix(1, 0) = Format(mdatBegin, "yyyy-MM-dd")
        .TextMatrix(1, 2) = "�ڳ����"
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = rsTemp("����")
            .TextMatrix(lngRow, 1) = rsTemp("NO")
            .TextMatrix(lngRow, 2) = IIf(IsNull(rsTemp("ժҪ")), "", rsTemp("ժҪ"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("��λ")), "", rsTemp("��λ"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("Ч��")), "", rsTemp("Ч��"))
            .TextMatrix(lngRow, 6) = Format(rsTemp("�ɹ�����"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 7) = Format(rsTemp("�ɹ���"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 8) = Format(rsTemp("Ӧ�����"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(rsTemp("�Ѹ����"), "###########0.00;-###########0.00; ; ")
                
            dblSum(1) = dblSum(1) + rsTemp("Ӧ�����")
            dblSum(2) = dblSum(2) + rsTemp("�Ѹ����")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .TextMatrix(lngRow, 0) = Format(mdatEnd, "yyyy-MM-dd")
        .TextMatrix(lngRow, 2) = "�ϼ�"
        .TextMatrix(lngRow, 8) = Format(dblSum(1), "###########0.00;-###########0.00; ; ")
        .TextMatrix(lngRow, 9) = Format(dblSum(2), "###########0.00;-###########0.00; ; ")
        .TextMatrix(lngRow, 10) = Format(dblBalance, "###########0.00;-###########0.00; ; ")
        
        
        Do Until lngRow = 1
            lngRow = lngRow - 1
            .TextMatrix(lngRow, 10) = Format(dblBalance, "###########0.00;-###########0.00; ; ")
            dblBalance = dblBalance + Val(.TextMatrix(lngRow, 9)) - Val(.TextMatrix(lngRow, 8))
        Loop
    End With

End Sub

Private Sub Fill�Ѹ��嵥()
'����:װ���Ѹ��嵥
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    
    '��ʼ�����
    ClearGrid mshDetail, 11
    With mshDetail
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .TextMatrix(0, 0) = "��ⵥ�ݺ�":   .colWidth(0) = 1000: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "���ݽ��":     .colWidth(1) = 1000: .ColAlignment(1) = 7
        .TextMatrix(0, 2) = "����ݺ�":   .colWidth(2) = 1000: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "����":         .colWidth(3) = 1100: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "ҩƷ����":     .colWidth(4) = 1500: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "���":         .colWidth(5) = 900: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "��λ":         .colWidth(6) = 900: .ColAlignment(6) = 1
        .TextMatrix(0, 7) = "����":         .colWidth(7) = 900: .ColAlignment(7) = 1
        .TextMatrix(0, 8) = "ʧЧ��":       .colWidth(8) = 1100: .ColAlignment(8) = 1
        .TextMatrix(0, 9) = "����":         .colWidth(9) = 1000: .ColAlignment(9) = 7
        .TextMatrix(0, 10) = "���":        .colWidth(10) = 1100: .ColAlignment(10) = 7
    End With
    '�õ���ѯ����
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '��ʼ��ѯ
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",C.���ﵥλ as ��λ,B.ʵ������/C.�����װ as ����"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",C.סԺ��λ as ��λ,B.ʵ������/C.סԺ��װ as ����"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",C.ҩ�ⵥλ as ��λ,B.ʵ������/C.ҩ���װ as ����"
    Else
        gstrSQL = ",C.�ۼ۵�λ as ��λ,B.ʵ������ as ����"
    End If
    '�õ��Ѹ��嵥
    gstrSQL = "select B.NO as ��ⵥ�ݺ�,b.���,E.NO as ����ݺ�,to_char(E.�������,'yyyy-MM-dd') as ����, " & _
              "         D.ͨ������ as ����,C.���,B.����,to_char(B.Ч��,'yyyy-MM-dd') as Ч��" & gstrSQL & ",A.��Ʊ��� as ��� " & _
              "    from ҩƷӦ����¼ A,ҩƷ�շ���¼ B,ҩƷĿ¼ C,ҩƷ��Ϣ D ,ҩƷ�����¼ E " & _
              "    Where A.�շ�ID = B.ID And B.ҩƷID = C.ҩƷID And C.ҩ��ID = D.ҩ��ID And A.������� = e.������� And e.��� = 1 " & _
              "        and E.�������>=to_date('" & strBegin & "','yyyyMMdd') and E.�������<to_date('" & strEnd & "','yyyyMMdd') and A.��ҩ��λID=" & lngID & _
              " and E.��λID= " & lngID & "  and E.��¼״̬<>2 order by B.NO,B.���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        mshDetail.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With mshDetail
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("��ⵥ�ݺ�")), " ", rsTemp("��ⵥ�ݺ�"))
            .TextMatrix(lngRow, 2) = rsTemp("����ݺ�")
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("����")), " ", rsTemp("����"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("���")), "", rsTemp("���"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rsTemp("��λ")), "", rsTemp("��λ"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, 8) = IIf(IsNull(rsTemp("Ч��")), "", rsTemp("Ч��"))
            .TextMatrix(lngRow, 9) = Format(rsTemp("����"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 10) = Format(rsTemp("���"), "###########0.00;-###########0.00; ; ")
                
            dblSum = dblSum + rsTemp("���")
            lngRow = lngRow + 1
            
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, 0) = "�ϼ�"
            .TextMatrix(lngRow, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 10) = Format(dblSum, "###########0.00;-###########0.00; ; ")
        End If
        '���㵥�ݵĺϼƽ��
        lngRow = 1
        Do While lngRow < .Rows - 1
            dblSum = Val(.TextMatrix(lngRow, 10))
            lngTemp = lngRow + 1
            
            Do While lngTemp < .Rows - 1
                If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) Then
                    dblSum = dblSum + Val(.TextMatrix(lngTemp, 10))
                Else
                    Exit Do
                End If
                lngTemp = lngTemp + 1
            Loop
            For lngCount = lngRow To lngTemp - 1
                .TextMatrix(lngCount, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            Next
            lngRow = lngTemp
        Loop
    End With
End Sub

Private Sub Fillδ���嵥()
'����:װ���Ѹ��嵥
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    
    '��ʼ�����
    ClearGrid mshDetail, 10
    With mshDetail
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .TextMatrix(0, 0) = "��ⵥ�ݺ�":   .colWidth(0) = 1000: .ColAlignment(0) = 1
        .TextMatrix(0, 1) = "���ݽ��":     .colWidth(1) = 1000: .ColAlignment(1) = 7
        .TextMatrix(0, 2) = "����":         .colWidth(2) = 1100: .ColAlignment(2) = 1
        .TextMatrix(0, 3) = "ҩƷ����":     .colWidth(3) = 1500: .ColAlignment(3) = 1
        .TextMatrix(0, 4) = "���":         .colWidth(4) = 900: .ColAlignment(4) = 1
        .TextMatrix(0, 5) = "��λ":         .colWidth(5) = 900: .ColAlignment(5) = 1
        .TextMatrix(0, 6) = "����":         .colWidth(6) = 900: .ColAlignment(6) = 1
        .TextMatrix(0, 7) = "ʧЧ��":         .colWidth(7) = 1100: .ColAlignment(7) = 1
        .TextMatrix(0, 8) = "����":         .colWidth(8) = 1000: .ColAlignment(8) = 7
        .TextMatrix(0, 9) = "���":         .colWidth(9) = 1100: .ColAlignment(9) = 7
    End With
    '�õ���ѯ����
    lngID = mshSum.RowData(mshSum.Row)
    If lngID = 0 Then Exit Sub
    
    '��ʼ��ѯ
    strBegin = Format(mdatBegin, "yyyyMMdd")
    strEnd = Format(mdatEnd + 1, "yyyyMMdd")
    
    '�õ�δ���嵥
    If mnuViewUnit(1).Checked = True Then
        gstrSQL = ",C.���ﵥλ as ��λ,B.ʵ������/C.�����װ as ����"
    ElseIf mnuViewUnit(2).Checked = True Then
        gstrSQL = ",C.סԺ��λ as ��λ,B.ʵ������/C.סԺ��װ as ����"
    ElseIf mnuViewUnit(3).Checked = True Then
        gstrSQL = ",C.ҩ�ⵥλ as ��λ,B.ʵ������/C.ҩ���װ as ����"
    Else
        gstrSQL = ",C.�ۼ۵�λ as ��λ,B.ʵ������ as ����"
    End If
    gstrSQL = "select B.NO as ��ⵥ�ݺ�,b.���,to_char(B.�������,'yyyy-MM-dd') as ����, " & _
              "         D.ͨ������ as ����,C.���,B.����,to_char(B.Ч��,'yyyy-MM-dd') as Ч��" & gstrSQL & ",A.��Ʊ��� as ��� " & _
              "    from ҩƷӦ����¼ A,ҩƷ�շ���¼ B,ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ�����¼ E  " & _
              "    Where A.�շ�ID = B.ID And B.ҩƷID = C.ҩƷID And C.ҩ��ID = D.ҩ��ID  And A.�������=E.�������(+) and E.���(+)=1 and E.������� is null " & _
              "        and B.�������>=to_date('" & strBegin & "','yyyyMMdd') and B.�������<to_date('" & strEnd & "','yyyyMMdd') and A.��ҩ��λID=" & lngID & _
              "  order by B.NO,B.���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        mshDetail.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With mshDetail
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("��ⵥ�ݺ�")), " ", rsTemp("��ⵥ�ݺ�"))
            .TextMatrix(lngRow, 2) = IIf(IsNull(rsTemp("����")), " ", rsTemp("����"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, 4) = IIf(IsNull(rsTemp("���")), "", rsTemp("���"))
            .TextMatrix(lngRow, 5) = IIf(IsNull(rsTemp("��λ")), "", rsTemp("��λ"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rsTemp("Ч��")), "", rsTemp("Ч��"))
            .TextMatrix(lngRow, 8) = Format(rsTemp("����"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(rsTemp("���"), "###########0.00;-###########0.00; ; ")
                
            dblSum = dblSum + rsTemp("���")
            lngRow = lngRow + 1
            
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, 0) = "�ϼ�"
            .TextMatrix(lngRow, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 9) = Format(dblSum, "###########0.00;-###########0.00; ; ")
        End If
        '���㵥�ݵĺϼƽ��
        lngRow = 1
        Do While lngRow < .Rows - 1
            dblSum = Val(.TextMatrix(lngRow, 9))
            lngTemp = lngRow + 1
            
            Do While lngTemp < .Rows - 1
                If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) Then
                    dblSum = dblSum + Val(.TextMatrix(lngTemp, 9))
                Else
                    Exit Do
                End If
                lngTemp = lngTemp + 1
            Loop
            For lngCount = lngRow To lngTemp - 1
                .TextMatrix(lngCount, 1) = Format(dblSum, "###########0.00;-###########0.00; ; ")
            Next
            lngRow = lngTemp
        Loop
    End With
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
    
    If ActiveControl Is mshSum Then
        blnPrint = Not (mshSum.Rows = 2 And mshSum.TextMatrix(1, 0) = "")
    Else
        blnPrint = Not (mshDetail.Rows = 2 And mshDetail.TextMatrix(1, 0) = "")
    End If

    mnuFilePreView.Enabled = blnPrint
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
    Call zlWebForum(Me.hwnd)
End Sub

