VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMain_���������Ŀ���� 
   Caption         =   "���ļ�Ŀ����"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14745
   Icon            =   "frmMain_���������Ŀ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   635
      SimpleText      =   $"frmMain_���������Ŀ����.frx":0E42
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain_���������Ŀ����.frx":0E89
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20929
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   14745
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   14625
         _ExtentX        =   25797
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "��ӡԤ��"
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
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "���ӱ������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ϴ�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�ϴ�HIS��Ŀ"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��ǰ�û���Ŀ"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split3"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
         Begin VB.TextBox txtFind 
            Height          =   315
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   4
            Top             =   195
            Width           =   1320
         End
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5595
      Top             =   75
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
            Picture         =   "frmMain_���������Ŀ����.frx":171D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":1937
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":1B51
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":1D6B
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":1F85
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":219F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":23B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":25D3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":27ED
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6390
      Top             =   15
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
            Picture         =   "frmMain_���������Ŀ����.frx":2A07
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":2C21
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":2E3B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":3055
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":326F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":3489
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":36A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":38BD
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_���������Ŀ����.frx":3AD7
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Bindings        =   "frmMain_���������Ŀ����.frx":3CF1
      Height          =   5295
      Left            =   60
      TabIndex        =   3
      Top             =   750
      Width           =   14580
      _cx             =   25717
      _cy             =   9340
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMain_���������Ŀ����.frx":3D06
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuDownloadCenter 
         Caption         =   "����������Ŀ(&A)"
         Enabled         =   0   'False
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpCenter 
         Caption         =   "�ϴ�HIS��Ŀ(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ����ǰ��¼(&D)"
         Shortcut        =   {DEL}
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
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
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
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "����(&A)"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "�ϴ�(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "ɾ��(&D)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain_���������Ŀ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsrue      As Integer
Private rsTemp          As ADODB.Recordset
Private mstrSort����ID  As String
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mintColumn As Integer
Dim mstrKey As String
Dim mint���� As Integer
Const conSql1 = "SELECT [ITEM_CLASS]" & vbNewLine & _
                "      ,[ITEM_CODE]" & vbNewLine & _
                "      ,[ITEM_NAME]" & vbNewLine & _
                "      ,[ITEM_SPECIFICATION]" & vbNewLine & _
                "      ,[UNIT]" & vbNewLine & _
                "      ,[STANDARD_PRICE]" & vbNewLine & _
                "      ,[ITEM_ON_DISPENSARY_RECEIPT]" & vbNewLine & _
                "      ,[ITEM_ON_RESIDENT_RECEIPT]" & vbNewLine & _
                "      ,[ITEM_NO_DEPT_STAT]" & vbNewLine & _
                "      ,[ITEM_NO_ACCOUNTANT_ITEM]" & vbNewLine & _
                "      ,[MEMO]" & vbNewLine & _
                "      ,[START_DATE]" & vbNewLine & _
                "      ,[STOP_DATE]" & vbNewLine & _
                "      ,[OPERATOR]" & vbNewLine & _
                "      ,[MODIFY_DATE]" & vbNewLine
Const conSql2 = "      ,[COLLATE_RELATION]" & vbNewLine & _
                "      ,[CONVERSION_RATE]" & vbNewLine & _
                "      ,[ITEM_FORM]" & vbNewLine & _
                "      ,[CHRONIC_CONVERSION_RATE]" & vbNewLine & _
                "      ,[CHRONIC_MIN_UNIT]" & vbNewLine & _
                "      ,[EXAMINE_PERSON]" & vbNewLine & _
                "      ,[EXAMINE_DATE]" & vbNewLine & _
                "      ,[EXAMINE_FLAG]" & vbNewLine & _
                "      ,[gkfldm]" & vbNewLine & _
                "      ,[kzyfdm]" & vbNewLine & _
                "      ,[zxks]" & vbNewLine & _
                "      ,[ypjcbz]" & vbNewLine & _
                "      ,[pydm]" & vbNewLine & _
                "      ,[zxksmc]" & vbNewLine & _
                "  FROM [PRICELIST_DICT]"
Const conSql = conSql1 & conSql2

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsrue = vNewValue
End Property

Private Sub Form_Load()
    Dim strField        As String
    Dim strFieldWIDth   As String
    Dim varField        As Variant
    Dim varFieldWIDth   As Variant
    Dim i               As Integer
                                 
    RestoreWinState Me, App.ProductName
    Call DataLoad
    If GetPersonSet Then
        RestoreFlexState vsfDetail, Me.Name
        'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
        strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", vsfDetail.Name & "����", "")
        strFieldWIDth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", vsfDetail.Name & "���", "")
        varField = Split(strField, ",")
        varFieldWIDth = Split(strFieldWIDth, ",")
        For i = 0 To UBound(varField)
            If varField(i) <> "" And Val(varFieldWIDth(i)) <> 0 Then
                If vsfDetail.ColIndex(varField(i)) <> -1 Then
                    vsfDetail.ColPosition(vsfDetail.ColIndex(varField(i))) = i
                    vsfDetail.ColWidth(i) = Val(varFieldWIDth(i))
                End If
            End If
        Next
        Me.WindowState = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", 0)
        If Me.WindowState = 0 Then
            Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left)
            Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top)
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    vsfDetail.Top = sngTop
    vsfDetail.Height = IIf(sngBottom - vsfDetail.Top > 0, sngBottom - vsfDetail.Top, 0)
    vsfDetail.Left = ScaleLeft

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    SaveFlexState vsfDetail, Me.Name
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", Me.WindowState
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top
End Sub

Private Sub mnuEditAdd_Click()
'    Dim str����ID       As String
'
'    With frmMain_���������Ŀ����༭
'        .Insure = mintInsrue
'        .Show vbModal
'        If Not .OkCancel Then
'            Set frmMain_���������Ŀ����༭ = Nothing
'            Exit Sub
'        End If
'        str����ID = .SickCode
'    End With
'    Set frmMain_���������Ŀ����༭ = Nothing
'    Call DataLoad
'    vsfSetRow vsfDetail, str����ID, "����ID"
'    If frmMain_���������Ŀ����༭.�༭����(lng����, "") = True Then
        '����¼�������Ѿ�������
'        Call SetMenu
'    End If
End Sub

Private Sub mnuEditModify_Click()
'    Dim str����ID As String
'    With frmMain_���������Ŀ����༭
'        If vsfDetail.Rows <= 1 Then Exit Sub
'        str����ID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����ID"))
'        .SickCode = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����ID"))
'        .Insure = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����"))
'        .Show vbModal
'        If Not .OkCancel Then
'            Set frmMain_���������Ŀ����༭ = Nothing
'            Exit Sub
'        End If
'    End With
'    Set frmMain_���������Ŀ����༭ = Nothing
'    Call DataLoad
'    vsfSetRow vsfDetail, str����ID, "����ID"
End Sub

Private Sub mnuEditDelete_Click()
'    Dim str����ID        As String
'    Dim strDelNote       As String
'
'    On Error GoTo errHandle
'    If vsfDetail.Rows <= 1 Then Exit Sub
'    str����ID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����ID"))
'    With frmCheckDelNote
'        .DelNote = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ȡ��ԭ��"))
'        .Show vbModal, Me
'        If (.DelNote = "") Then
'            Set frmCheckDelNote = Nothing
'            Exit Sub
'        End If
'        strDelNote = .DelNote
'    End With
'    Set frmCheckDelNote = Nothing
'    gstrSQL = "dl_����_�����޶�_Cancel(" & vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����")) & ",'" & str����ID & "','" & UserInfo.���� & "','" & strDelNote & "')"
'    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'    Call DataLoad
'    vsfSetRow vsfDetail, str����ID, "����ID"
'
'    Call SetMenu
'    MousePointer = vbDefault
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'    MousePointer = vbDefault
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfDetail Is Nothing Then Exit Sub
    LockWindowUpdate 0
    '���ô�ӡ��������
    Set objPrint.Body = vsfDetail
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ�ˣ�" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ�䣺" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuEditAdd_Click
        Case 1
            mnuEditModify_Click
        Case 2
            mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuViewFind_Click()
    Call DataLoad
    txtFind.Text = ""
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
'    For i = 0 To 3
'        mnuViewIcon(i).Checked = False
'    Next
'    mnuViewIcon(Index).Checked = True
End Sub

Private Sub mnuViewRefresh_Click()
    Call DataLoad
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
    Dim lngCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For lngCount = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(lngCount).Caption = IIf(mnuViewToolText.Checked = True, tbrThis.Buttons(lngCount).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    cbrThis.Refresh
    Call Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Modify"
            mnuEditModify_Click
        Case "View"
'            If lvwItem.View = 3 Then
'                mnuViewIcon(0).Checked = True
'                lvwItem.View = 0
'            Else
'                mnuViewIcon(lvwItem.View + 1).Checked = True
'                lvwItem.View = lvwItem.View + 1
'            End If
        Case "Find"
            mnuViewFind_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
'    For i = 0 To 3
'        mnuViewIcon(i).Checked = False
'    Next
'    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
'    lvwItem.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
    
End Sub

Private Sub SetMenu()
'���ܣ����ݵ�ǰ�������ò˵��Ŀ�����
    Dim bln���Թ� As Boolean
    Dim bln������ As Boolean
    
'    Call FillItem
'    stbThis.Panels(2).Text = lvwKind_S.SelectedItem.Text & "����" & lvwItem.ListItems.Count & "�����ּ�¼"
    
    tbrThis.Buttons("New").Enabled = True
    mnuEdit.Enabled = True
'    mnuEditAdd.Enabled = True
    mnuShortMenu(0).Enabled = True
    
    If vsfDetail.Rows > 1 Then
        tbrThis.Buttons("Modify").Enabled = True
        tbrThis.Buttons("Delete").Enabled = True
        tbrThis.Buttons("Split1").Enabled = True
'        mnuEditModify.Enabled = True
        mnuShortMenu(1).Enabled = True
        mnuShortMenu(2).Enabled = True
    Else
        tbrThis.Buttons("Modify").Enabled = False
        tbrThis.Buttons("Delete").Enabled = False
        tbrThis.Buttons("Split1").Enabled = False
'        mnuEditModify.Enabled = False
        mnuShortMenu(1).Enabled = False
        mnuShortMenu(2).Enabled = False
    End If
End Sub

Private Sub DataLoad()
    gstrSQL = conSql
    gstrSQL = gstrSQL & " And A.����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mintInsrue)
    Set vsfDetail.DataSource = rsTemp
    Call vsfDetail_RowColChange
    Call SetMenu
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If Len(txtFind.Text) > 1 And KeyAscii = 13 Then
        vsfSetRow vsfDetail, txtFind.Text, "���ֱ���,��������,���ּ���"
    End If
End Sub

'==============================================================================
'=���ܣ� �����λ��¼ vsfDetail
'==============================================================================
Private Sub vsfDetail_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
'    vsfSetRow vsfDetail, mstrSort����ID, "����ID"
    lngRow = vsfDetail.FindRow(mstrSort����ID, -1, vsfDetail.ColIndex("����ID"), False, True)
    If lngRow > 0 Then vsfDetail.Row = lngRow
    vsfDetail.ShowCell lngRow, 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Cancel = True
End Sub

'==============================================================================
'=���ܣ� ĳ�в����ƶ�λ�� vsfDetail[ͼ��]
'==============================================================================
Private Sub vsfDetail_BeforeMoveColumn(ByVal COL As Long, Position As Long)
    If COL = vsfDetail.ColIndex("ͼ��") Then
        Position = -1
    Else
        If Position <= vsfDetail.ColIndex("ͼ��") Then Position = COL
    End If
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼����ID vsfDetail
'==============================================================================
Private Sub vsfDetail_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSort����ID = "" & vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С vsfDetail[ͼ��]
'==============================================================================
Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If COL = vsfDetail.ColIndex("ͼ��") Then Cancel = True
End Sub

'==============================================================================
'=���ܣ� ˫������޸Ĺ��� vsfDetail
'==============================================================================
Private Sub vsfDetail_DblClick()
    On Error GoTo ErrH
    If vsfDetail.MouseRow <= 0 Then Exit Sub
    mnuEditModify_Click
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� vsfDetail
'==============================================================================
Private Sub vsfDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
            PopupMenu mnuShort
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=���ܣ����б任ʱ
'==============================================================================
Private Sub vsfDetail_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant
    On Error GoTo ErrH
    DoEvents
    If vsfDetail.Rows = 1 Then
        stbThis.Panels(2) = "��ǰ��ʾ�� 0 ����Ŀ��"
        Exit Sub
    End If
    If vsfDetail.ColIndex("����ID") <= 0 Then Exit Sub
    stbThis.Panels(2) = "��ǰ��ʾ�� " & vsfDetail.Rows - 1 & " ����Ŀ��"
    '������ϸ����
    gstrSQL = "select C.����ID ,A.ID as �շ�ID,���� as �շѱ���,���� as �շ�����,B.���� as �շѽ��" & vbCrLf & _
                "from �շ�ϸĿ A,(Select x.�շ�ϸĿid, y.�ּ� As ���� from" & vbCrLf & _
                "(Select �շ�ϸĿid, Max(ID) As ID" & vbCrLf & _
                "From �շѼ�Ŀ" & vbCrLf & _
                "Where Sysdate >= ִ������ And Sysdate <= ��ֹ����" & vbCrLf & _
                "Group By �շ�ϸĿid) X, �շѼ�Ŀ Y" & vbCrLf & _
                "Where x.Id = y.Id) B,����_���ֲ��� C" & vbCrLf & _
                "where ���='M' And a.ID = B.�շ�ϸĿid And A.ID=C.�շ�ID AND C.����=[1] AND C.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����")), vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����ID")))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfProject_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Cancel = True
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� vsfDetail
'==============================================================================
Private Sub vsfProject_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
            PopupMenu mnuShort
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
