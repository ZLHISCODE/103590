VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMedicalStationRequest 
   Caption         =   "������뵥"
   ClientHeight    =   7800
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11820
   Icon            =   "frmMedicalStationRequest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11820
   Begin VB.Frame fra 
      Caption         =   "ִ�п���"
      Height          =   1875
      Left            =   105
      TabIndex        =   13
      Top             =   5580
      Width           =   2565
      Begin VB.ListBox lstDept 
         Height          =   1530
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3570
      Left            =   4140
      TabIndex        =   0
      Top             =   2025
      Width           =   7170
      _cx             =   12647
      _cy             =   6297
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
      Cols            =   6
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
      MergeCells      =   1
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
         X1              =   270
         X2              =   270
         Y1              =   435
         Y2              =   1650
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8370
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":020A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":05A4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":083A
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":0BD4
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":0F6E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":1308
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":16A2
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":1938
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":1BCE
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":1E64
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":20FA
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationRequest.frx":2390
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15769
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":2C24
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":339E
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":3B18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":3D32
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":3F4C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":416C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":438C
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":4B06
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":5280
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":59FA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":5C14
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":5E2E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":604E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":626E
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11820
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   11700
         _ExtentX        =   20638
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&V.Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��(Alt+V)"
               Object.Tag             =   "&V.Ԥ��"
               ImageKey        =   "PrintView"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&P.��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ(Alt+P)"
               Object.Tag             =   "&P.��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra3 
      Height          =   630
      Left            =   4110
      TabIndex        =   4
      Top             =   5985
      Width           =   3315
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationRequest.frx":69E8
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Tag             =   "����"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.����"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   1
         Tag             =   "����"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Frame fra1 
      Height          =   3090
      Left            =   105
      TabIndex        =   9
      Top             =   720
      Width           =   2580
      Begin VB.CheckBox chk 
         Caption         =   "��������Ŀ(&2)"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   435
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������Ŀ(&1)"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.ListBox lstStyle 
         Height          =   2370
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   660
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7155
      Top             =   975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":6C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":6F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":7DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":D5CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":DFDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRequest.frx":13AC8
            Key             =   "Query"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "ȫ��(&C)"
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
Attribute VB_Name = "frmMedicalStationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mlng����id As Long
Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
        
    If vData = False Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
'
    
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng����id As Long = 0, Optional ByVal bln�������뵥 As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mlng����id = lng����id
    mlngKey = lngKey
    
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng����id) = False Then Exit Function
    
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng����id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
            
    
    If mlng����id > 0 Then
        
        gstrSQL = "Select a.����id As ID,0 As ѡ��,b.����,b.�����,b.������,b.���֤��,a.����� " & _
                    "From �����Ա���� a, " & _
                        "������Ϣ b " & _
                    "Where b.����id=a.����id " & _
                            "And a.�Ǽ�id=[1] And a.����id=[2]"
                            
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, mlng����id)
        
    Else
    
        gstrSQL = "Select a.����id As ID,0 As ѡ��,b.����,b.�����,b.������,b.���￨��,b.���֤��,a.����� " & _
                    "From �����Ա���� a, " & _
                        "������Ϣ b " & _
                    "Where b.����id=a.����id " & _
                            "And a.�Ǽ�id=[1] Order by b.����� "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        
    End If
    If rs.BOF = False Then
    
        vsf.TextMatrix(0, GetCol(vsf, "ѡ��")) = "ѡ��"
        Call LoadGrid(vsf, rs, , , ils13)
        Call AppendRows(vsf, lnX, lnY)
        vsf.TextMatrix(0, GetCol(vsf, "ѡ��")) = ""
        
    End If
    
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    Dim strSectoin As String
    Dim strTmp As String
    Dim lngLoop As Long
    
    strSectoin = "˽��ģ��\" & App.ProductName & "\������뵥����"
    
    On Error GoTo errHand
    
    vsf.MergeCells = flexMergeFree
        
       
    strVsf = ",240,1,1,1,ѡ��;����,750,1,1,1,;�����,810,7,1,1,;������,810,7,1,1,;���￨��,0,1,1,0,;�����,990,1,1,1,;���֤��,1200,1,1,0,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(GetCol(vsf, "ѡ��")) = flexDTBoolean
    vsf.Editable = flexEDKbdMouse

    Call AppendRows(vsf, lnX, lnY)

    lstStyle.Clear
    gstrSQL = "select * from ���Ƽ���걾 Order By  ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            lstStyle.AddItem zlCommFun.NVL(rs("����"))
            rs.MoveNext
        Loop
    End If
    
    
    strTmp = GetSetting("ZLSOFT", strSectoin, "��ӡִ�п���", "")
    If strTmp <> "" Then strTmp = "'" & strTmp & "'"
    lstDept.Clear
    gstrSQL = "Select ID,���� From ���ű� Where ID In (Select ����id From ��������˵�� Where �������� In ('���','����'))"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            lstDept.AddItem zlCommFun.NVL(rs("����").Value)
            lstDept.ItemData(lstDept.NewIndex) = zlCommFun.NVL(rs("ID").Value, 0)
            If strTmp = "" Then
                lstDept.Selected(lstDept.NewIndex) = True
            Else
                If InStr(strTmp, "'" & zlCommFun.NVL(rs("ID").Value, 0) & "'") > 0 Then
                    lstDept.Selected(lstDept.NewIndex) = True
                End If
            End If
            rs.MoveNext
        Loop
    End If
    
        
    strTmp = GetSetting("ZLSOFT", strSectoin, "����1", "")
    
    If Trim(strTmp) <> "" And InStr(strTmp, "'") > 0 Then
        Call RestoreCondition(Mid(strTmp, InStr(strTmp, "'") + 1))
    End If

    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal blnGroup As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    Dim lng����id As Long
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    Dim strDeptID As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    If chk(1).Value = 0 And chk(0).Value = 0 Then Exit Function
    
    strDeptID = ""
    For lngLoop = 0 To lstDept.ListCount - 1
        If lstDept.Selected(lngLoop) Then
            strDeptID = strDeptID & "," & lstDept.ItemData(lngLoop)
        End If
    Next
    If strDeptID <> "" Then strDeptID = strDeptID & ","
    
    intѡ�� = GetCol(vsf, "ѡ��")
    strTmp = ""
    
    For lngLoop = 0 To lstStyle.ListCount - 1
        If lstStyle.Selected(lngLoop) Then
            strTmp = strTmp & "''" & lstStyle.List(lngLoop)
        End If
    Next
    
    If strTmp <> "" Then strTmp = strTmp & "''"
    
    For lngLoop = 1 To vsf.Rows - 1
    
        If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, intѡ��))) = 1 Then
                
            Call OutPutQuestBill(Me, mlngKey, Val(vsf.RowData(lngLoop)), strDeptID, strTmp, (chk(1).Value = 1), (chk(0).Value = 1), bytMode)
            
            '�����Ԥ����ֻһ��Ԥ��
            If bytMode = 1 Then Exit For
            
        End If
    Next
      
    PrintData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then
        Resume
    End If

End Function



Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 8)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyM
            If tbrThis.Buttons("�ʼ�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�ʼ�"))
        Case vbKeyV
            If tbrThis.Buttons("Ԥ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("Ԥ��"))
        Case vbKeyP
            If tbrThis.Buttons("��ӡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("��ӡ"))
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

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()

    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra1
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
    End With
    
    With lstStyle
        .Height = fra1.Height - .Top - 60
    End With

'    With fra4
'        .Left = fra1.Left
'        .Top = fra1.Top + fra1.Height + 15
'        .Width = fra1.Width
'    End With
'    tbr.Top = lvw.Top + lvw.Height + 30
    
    With fra
        .Left = fra1.Left
        .Top = fra1.Top + fra1.Height + 15
        .Width = fra1.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    lstDept.Move lstDept.Left, lstDept.Top, lstDept.Width, fra.Height - lstDept.Top - 45
    
    With vsf
        .Left = fra1.Left + fra1.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fra3.Height
    End With
    
    With fra3
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height
        .Width = vsf.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveWinState(Me, App.ProductName)
    
    Dim strTmp As String
    Dim lngLoop As Long
    Dim strSectoin As String
    
    '�����˳��������ĸ���
    strSectoin = "˽��ģ��\" & App.ProductName & "\������뵥����"
    
    On Error Resume Next '���û�иü�ֵ���ͻ����
    DeleteSetting "ZLSOFT", strSectoin 'ɾ����ǰ������
    On Error GoTo 0
    
    Call SaveSetting("ZLSOFT", strSectoin, "����1", "'" & SaveCondition)
    
    strTmp = ""
    For lngLoop = 0 To lstDept.ListCount - 1
        If lstDept.Selected(lngLoop) Then
            strTmp = strTmp & "'" & lstDept.ItemData(lngLoop)
        End If
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    Call SaveSetting("ZLSOFT", strSectoin, "��ӡִ�п���", strTmp)
    
End Sub

Private Sub RestoreCondition(ByVal strTag As String)
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    If strTag = "" Then Exit Sub
    
    varTmp = Split(strTag, "'")
    
    On Error Resume Next
    
    chk(0).Value = Val(varTmp(0))
    chk(1).Value = Val(varTmp(1))
    
    For lngLoop = 2 To UBound(varTmp) - 1
        strTmp = strTmp & "'" & varTmp(lngLoop)
    Next
    strTmp = strTmp & "'"
    
    For lngLoop = 0 To lstStyle.ListCount - 1
    
        '
        If InStr(strTmp, "'" & lstStyle.List(lngLoop) & "'") > 0 Then
            lstStyle.Selected(lngLoop) = True
        Else
            lstStyle.Selected(lngLoop) = False
        End If
        
    Next

End Sub

Private Function SaveCondition() As String
    '
    Dim strTmp As String
    Dim lngLoop As Long
    
    strTmp = chk(0).Value & "'" & chk(1).Value
    For lngLoop = 0 To lstStyle.ListCount - 1
        If lstStyle.Selected(lngLoop) Then
            strTmp = strTmp & "'" & lstStyle.List(lngLoop)
        End If
    Next
    strTmp = strTmp & "'0"
    SaveCondition = strTmp
    
End Function

Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    
    intѡ�� = GetCol(vsf, "ѡ��")
    If intѡ�� >= 0 Then
    
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, intѡ��) = 0
            End If
        Next
        
        EditChanged = False
        
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    
    Call PrintData(2)

End Sub

Private Sub mnuFilePrintView_Click()
    
    Call PrintData(1)
    
End Sub


Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    
    intѡ�� = GetCol(vsf, "ѡ��")
    If intѡ�� >= 0 Then
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, intѡ��) = 1
                EditChanged = True
            End If
        Next
    End If
    
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

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu

    Case 3
        
        mobjPopMenu.Add 1, "&1.����", , , True, , (lbl(1).Tag = "����")
        mobjPopMenu.Add 2, "&2.�����", , , True, , (lbl(1).Tag = "�����")
        mobjPopMenu.Add 3, "&3.������", , , True, , (lbl(1).Tag = "������")
        mobjPopMenu.Add 4, "&4.���￨��", , , True, , (lbl(1).Tag = "���￨��")
        mobjPopMenu.Add 5, "&5.����ƴ��", , , True, , (lbl(1).Tag = "����ƴ��")
        mobjPopMenu.Add 6, "&6.�������", , , True, , (lbl(1).Tag = "�������")
        mobjPopMenu.Add 7, "&7.���֤��", , , True, , (lbl(1).Tag = "���֤��")
        mobjPopMenu.Add 8, "&8.�����", , , True, , (lbl(1).Tag = "�����")
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu

    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "�ʼ�"
        
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    Dim blnCard As Boolean
    Dim lngRow As Long
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    strCol = Mid(lbl(1).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "���￨��" And KeyAscii <> vbKeyReturn Then
        '���￨�ţ��Զ�ʶ��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn
        End If

    End If
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            strCol = Mid(lbl(1).Caption, 4)
            Select Case strCol
            Case "����ƴ��"
                lngCol = GetCol(vsf, "����")
            Case "�������"
                lngCol = GetCol(vsf, "����")
            Case Else
                lngCol = GetCol(vsf, strCol)
            End Select
            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf.Row + 1 <= vsf.Rows - 1 Then
                For lngLoop = vsf.Row + 1 To vsf.Rows - 1
                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf.Row
                    lngRow = 0
                    Select Case strCol
                    Case "�����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "������"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���￨��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "���֤��"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "����ƴ��"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "�������"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "û���ҵ�����Ҫ�����Ϣ��"
                txt(Index).Text = ""
            Else
                vsf.ShowCell lngRow, vsf.Col
                vsf.Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    Dim intѡ�� As Integer
    
    intѡ�� = GetCol(vsf, "ѡ��")
    
    If intѡ�� >= 0 Then
        If Abs(Val(vsf.TextMatrix(Row, intѡ��))) = 1 Then
            EditChanged = True
            Exit Sub
        End If
            
        For lngLoop = 1 To vsf.Rows - 1
            If Abs(Val(vsf.TextMatrix(lngLoop, intѡ��))) = 1 Then
                EditChanged = True
                Exit Sub
            End If
        Next
        
        If lngLoop = vsf.Rows Then EditChanged = False
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> GetCol(vsf, "ѡ��") Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
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

