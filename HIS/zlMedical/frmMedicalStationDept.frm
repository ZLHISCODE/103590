VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationDept 
   Caption         =   "ִ�п��ҵ���"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmMedicalStationDept.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10350
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   90
      ScaleHeight     =   435
      ScaleWidth      =   5565
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1305
      Width           =   5565
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   90
         Width           =   3720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.�����Ŀ"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   1
      Left            =   345
      ScaleHeight     =   435
      ScaleWidth      =   5565
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4395
      Width           =   5565
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   105
         Width           =   2280
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.����Ϊ��ִ�п���"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   0
         Top             =   150
         Width           =   1620
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10350
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   10230
         _ExtentX        =   18045
         _ExtentY        =   1270
         ButtonWidth     =   1402
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&D.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+D)"
               Object.Tag             =   "&D.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   4470
      Top             =   750
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
            Picture         =   "frmMedicalStationDept.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":28E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":2B00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   5460
      Top             =   750
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
            Picture         =   "frmMedicalStationDept.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":3C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":438E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":45AE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5865
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationDept.frx":47CE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1980
      Left            =   165
      TabIndex        =   1
      Top             =   1905
      Width           =   3795
      _cx             =   6694
      _cy             =   3492
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      ExplorerBar     =   0
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8430
      Top             =   3855
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
            Picture         =   "frmMedicalStationDept.frx":5062
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":53FC
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":5796
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4875
      Left            =   7020
      TabIndex        =   10
      Top             =   855
      Width           =   3285
      Begin VB.OptionButton opt 
         Caption         =   "�ҵ�����Ϊδѡ��(&5)"
         Height          =   210
         Index           =   1
         Left            =   1005
         TabIndex        =   20
         Top             =   2325
         Width           =   2205
      End
      Begin VB.OptionButton opt 
         Caption         =   "�ҵ�����Ϊѡ��(&4)"
         Height          =   210
         Index           =   0
         Left            =   1005
         TabIndex        =   19
         Top             =   1965
         Value           =   -1  'True
         Width           =   2205
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   555
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&3.��������λ"
         Height          =   240
         Left            =   4710
         TabIndex        =   14
         Top             =   210
         Width           =   1425
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1125
         TabIndex        =   13
         Top             =   1365
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         TabIndex        =   12
         Top             =   930
         Width           =   1995
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1110
         TabIndex        =   11
         Text            =   "cbo"
         Top             =   195
         Width           =   1995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.������λ"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.��    ��"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�� �� ��"
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Tag             =   "�����"
         Top             =   1005
         Width           =   900
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "ȫ��(&C)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFile_1 
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
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
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStationDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mstrSQL As String
Private mblnChangeEdit As Boolean

Private mlngKey As Long

Private Enum mCol
    ѡ�� = 0
    �����
    ����
    ִ�п���
    ������λ
    ���
    ����id
    ִ�п���id
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    mnuFileSave.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
    End If

        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled

    
End Property

Private Sub RefreshState()
    
    Dim lngLoop As Long
    Dim intCount As Integer
    
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, 0))) = 1 Then
            intCount = intCount + 1
        End If
    Next
    
    stbThis.Panels(2).Text = "��ǰѡ�� " & intCount & " ��"
End Sub

Private Function SearchSelect(ByVal blnSel As Boolean) As Boolean
    '==================================================================================================================
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind1 As Boolean
    Dim blnFind2 As Boolean
    Dim blnFind3 As Boolean
    Dim blnFind4 As Boolean
    Dim lngStartRow As Long

    lngStartRow = vsf.Row
    For lngRow = lngStartRow To vsf.Rows - 1
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(2).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.������λ)), UCase(cbo(2).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.�����)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
        End If
        
        If cbo(3).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.���)), UCase(cbo(3).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            '�ҵ�

            vsf.TextMatrix(lngRow, mCol.ѡ��) = IIf(blnSel, 1, 0)
            SearchSelect = True
            vsf.Row = lngRow

        End If
    Next
    
    For lngRow = 1 To lngStartRow
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.������λ)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.�����)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
        End If
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.���)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
        
            vsf.TextMatrix(lngRow, mCol.ѡ��) = IIf(blnSel, 1, 0)
            SearchSelect = True
            
            vsf.Row = lngRow
            
        End If
    Next
    
    If SearchSelect Then
        vsf.ShowCell vsf.Row, vsf.Col
        vsf.SetFocus
    End If
    
End Function

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '����:�����޸�״̬���ð�ť���˵��ȵĿ���״̬
    '-----------------------------------------------------------------------------------------
    
    mnuFileSave.Enabled = True
        
    If mblnChangeEdit = False Then mnuFileSave.Enabled = False
        
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    If vsf.Rows = 2 And Trim(vsf.TextMatrix(1, 1)) = "" Then
        stbThis.Panels(2).Text = "û����Ϣ��"
    Else
        stbThis.Panels(2).Text = "���ҵ� " & vsf.Rows - 1 & " ����Ϣ��"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ���༭����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    mblnStartUp = False
    
    Call cbo_Click(0)
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "ѡ��,540,4,1,1,;�����,900,7,1,1,;����,900,1,1,1,;ִ�п���,1200,1,1,1,;������λ,1500,1,1,1,;���,1500,1,1,1,;����id,0,1,1,0,;ִ�п���id,0,1,1,0,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.ѡ��) = flexDTBoolean
    vsf.Editable = flexEDKbdMouse

    '��ȡ��Ŀ�嵥
    gstrSQL = "Select Distinct b.����,b.ID from �����Ŀ�嵥 a,������ĿĿ¼ b where a.������Ŀid=b.id and a.�Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    Call AddComboData(cbo(0), rs)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    '��ȡҽ������
    gstrSQL = "Select Distinct b.����,b.ID from ���ű� b where b.id in (select ����id From ��������˵�� Where �������� In ('���','����'))"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Call AddComboData(cbo(1), rs)
    If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0


    gstrSQL = "SELECT Distinct B.������λ " & _
                "FROM �����Ա���� A,������Ϣ B " & _
                "WHERE B.������λ Is Not Null And A.���״̬ IN (1,4) AND A.��챨�� In ([2],[3]) AND A.����id=B.����id and A.�Ǽ�id=[1]"
                
    cbo(2).Clear
    cbo(2).AddItem ""
    cbo(2).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 0, 1)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(2).AddItem rs("������λ").Value
            rs.MoveNext
        Loop
    End If
    
    gstrSQL = "SELECT Distinct A.������� " & _
                "FROM �����Ա���� A " & _
                "WHERE A.������� Is Not Null And A.���״̬ IN (1,4) AND A.�Ǽ�id=[1]"
                    
    cbo(3).Clear
    cbo(3).AddItem ""
    cbo(3).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(3).AddItem rs("�������").Value
            rs.MoveNext
        Loop
    End If
    
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Call ResetVsf(vsf)
    Call AppendRows(vsf, lnX, lnY)
    
    mstrSQL = "Select 0 As ѡ��,y.����id,y.����id AS ID,y.����,y.�����,y.������,y.������λ,x.����id,x.ִ�п���id,z.���� As ִ�п���,x.������� As ��� From " & _
                "( " & _
                "Select Distinct  b.����id,f.����id,a.ִ�п���id,b.������� " & _
                "from �����Ŀҽ�� a,�����Ա���� b,�����Ŀ�嵥 c,���ǼǼ�¼ d,����ҽ����¼ e,����ҽ������ f " & _
                "Where c.������Ŀid = [1] and c.�Ǽ�id=[2] " & _
                      "and a.�嵥id=c.id " & _
                      "and c.�Ǽ�id=b.�Ǽ�id " & _
                      "and a.����id=b.����id " & _
                      "and d.id=b.�Ǽ�id " & _
                      "and e.�Һŵ�=d.���� " & _
                      "and e.������Դ=4 and b.��챨��=1 And b.���״̬=4 " & _
                      "and e.ҽ��״̬<>4 " & _
                      "and e.������� In ('C','D') and e.������Ŀid=c.������Ŀid and e.����id=b.����id " & _
                      "and f.ҽ��id(+)=e.id " & _
                ") x,������Ϣ y,���ű� z " & _
                "where x.����id=y.����id and z.id=x.ִ�п���id and x.����id Is Null"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, mlngKey)
    If rs.BOF = False Then
        Call LoadGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    '�����������ֵ�Ƿ���ȷ,��Ҫ�Ǽ��鹫ʽ
    
    On Error GoTo errHand
    
    
            
    ValidData = True
    
    Exit Function
errHand:
    LocationObj txt(1)
    strError = "�������ֵ��ʽ���Ϸ���"
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim str�ɼ�No As String
    Dim strNo As String
    Dim lngSendNo As Long
    Dim lngDept As Long
    Dim lngTotal As Long
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    Me.Enabled = False
    Call frmWait.OpenWait(Me, "����ִ�еص�")
    frmWait.WaitInfo = "���ڵ���ִ�еص�..."
    
    lngSendNo = GetNextNo(10)
    lngDept = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
    
    gstrSQL = "Select a.*,b.����id As ���� From �����Ŀ�嵥 a,�����Ŀҽ�� b Where a.�Ǽ�id=[1] and a.������Ŀid=[2] And a.id=b.�嵥id and b.ִ�п���id<> [3] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, cbo(0).ItemData(cbo(0).ListIndex), cbo(1).ItemData(cbo(1).ListIndex))
            
    lngTotal = vsf.Rows - 1
    For mlngLoop = 1 To lngTotal
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 And Val(vsf.RowData(mlngLoop)) > 0 Then

            frmWait.WaitInfo = "���ڵ���ִ�еص㡰" & vsf.TextMatrix(mlngLoop, mCol.����) & " ��..." & Format(100 * mlngLoop / lngTotal, "0.00") & "%"
            
            rs.Filter = ""
            rs.Filter = "����=" & Val(vsf.RowData(mlngLoop))
            If rs.RecordCount > 0 Then
                
                Call SQLRecord(rsSQL)
                
                strSQL = "zl_�����Ŀҽ��_Modify(" & mlngKey & "," & cbo(0).ItemData(cbo(0).ListIndex) & "," & Val(vsf.RowData(mlngLoop)) & "," & cbo(1).ItemData(cbo(1).ListIndex) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
                '�������õ��ݺ�
                str�ɼ�No = ""
                strNo = ""
                If Val(zlCommFun.NVL(rs("����;��").Value, 1)) = 1 Then
                    '����
                    strNo = GetNextNo(14)
                Else
                    strNo = GetNextNo(13)
                End If
                
                If Val(zlCommFun.NVL(rs("�ɼ���ʽid").Value, 0)) > 0 Then
                    '�ɼ�
                    If Val(zlCommFun.NVL(rs("����;��").Value, 1)) = 1 Then
                        '����
                        str�ɼ�No = GetNextNo(14)
                    Else
                        str�ɼ�No = GetNextNo(13)
                    End If
                End If
                
                
                strSQL = "ZL_�����Ŀҽ��_NO(" & rs("ID").Value & "," & Val(vsf.RowData(mlngLoop)) & ",'" & strNo & "','" & str�ɼ�No & "')"
'                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Call SQLRecordAdd(rsSQL, strSQL)
                
                strSQL = "zl_�����Ա����_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(mlngLoop)) & "," & lngDept & "," & rs("ID").Value & ",1)"
'                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Call SQLRecordAdd(rsSQL, strSQL)
                
                blnTran = True
                gcnOracle.BeginTrans
                
                If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
                For lngCount = 1 To rsSQL.RecordCount
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                    rsSQL.MoveNext
                Next
                
                
                '������ط���
                If MakeMedicalCharge(rsSQL, mlngKey) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If
                
                strSQL = "zl_�����Ա����_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(mlngLoop)) & "," & lngDept & "," & rs("ID").Value & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
            End If
        End If
    Next
    
    frmWait.CloseWait
    Me.Enabled = True
    
    SaveData = True

    Exit Function

errHand:
    frmWait.CloseWait
    Me.Enabled = True
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then
        gcnOracle.RollbackTrans
        ShowSimpleMsg "δ��ȫ�����ɹ��򲿷ݵ����ɹ���"
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    If mblnStartUp Then Exit Sub
    
    If Index = 0 Then
        If cbo(Index).ListIndex >= 0 Then
            Call ReadData(cbo(Index).ItemData(cbo(Index).ListIndex))
        End If
    ElseIf Index > 1 Then

    End If
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdSelect_Click()
    If SearchSelect(opt(0).Value) Then

        EditChanged = True
        Call RefreshState
        
    End If
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '���鲿��
'    cbo(1).Clear
'    For mlngLoop = 0 To mfrmMain.cboDept.ListCount - 1
'        cbo(1).AddItem mfrmMain.cboDept.List(mlngLoop)
'        cbo(1).ItemData(cbo(1).NewIndex) = mfrmMain.cboDept.ItemData(mlngLoop)
'    Next
'    cbo(1).ListIndex = mfrmMain.cboDept.ListIndex
'
'
'    If cbo(1).ListIndex = -1 Then
'        zlControl.CboLocate cbo(1), UserInfo.����ID, True
'        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
'    End If
    
'    txt(0).Text = ""
'    txt(2).Text = ""

    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyB
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBack(0).Move 0, IIf(cbrThis.Visible, cbrThis.Height, 0), Me.ScaleWidth - fra.Width
    vsf.Move 0, picBack(0).Top + picBack(0).Height, picBack(0).Width, Me.ScaleHeight - picBack(0).Top - picBack(0).Height - IIf(stbThis.Visible, stbThis.Height, 0) - picBack(1).Height
    picBack(1).Move 0, vsf.Top + vsf.Height, vsf.Width
    
    fra.Move vsf.Left + vsf.Width, picBack(0).Top - 90, fra.Width, Me.ScaleHeight - fra.Top - IIf(stbThis.Visible, stbThis.Height, 0) + 90

    Call AppendRows(vsf, lnX, lnY)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChangeEdit Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuFileClearAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 0
    Next
    mblnChangeEdit = False
    Call AdjustEnableState
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    If mblnChangeEdit Then
        
        If MsgBox("���Ҫ��ѡ�е���Ŀ����Ϊ��ִ�п�����", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        If SaveData() = False Then Exit Sub
        
        mblnOK = True
        
        mblnChangeEdit = False
        Call AdjustEnableState

        ShowSimpleMsg "���ִ�п��ҵ����ɹ���"
        
        Call ResetVsf(vsf)
        txt(1).Text = ""
        Call AppendRows(vsf, lnX, lnY)
        
        Exit Sub
    End If
End Sub

Private Sub mnuFileSelectAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 Then
            vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 1
        End If
    Next
    
    mblnChangeEdit = True
    Call AdjustEnableState
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
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
    Dim intLoop As Integer

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub


Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    cmdSelect.Default = True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    End Select
    
    cmdSelect.Default = False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.ѡ��))) = 1 Then
        mblnChangeEdit = True
        Call AdjustEnableState
        Exit Sub
    End If
        
    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 Then
            mblnChangeEdit = True
            Call AdjustEnableState
            Exit Sub
        End If
    Next
    
    If mlngLoop = vsf.Rows Then
        mblnChangeEdit = False
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

