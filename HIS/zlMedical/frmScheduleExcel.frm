VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmScheduleExcel 
   Caption         =   "���ɵǼǱ��"
   ClientHeight    =   5460
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9720
   Icon            =   "frmScheduleExcel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9720
   Begin VB.TextBox txtInfo 
      Height          =   2295
      Left            =   11700
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "frmScheduleExcel.frx":076A
      Top             =   5625
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   3960
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHead 
      Height          =   2295
      Left            =   11325
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "frmScheduleExcel.frx":0F62
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   5100
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmScheduleExcel.frx":175A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9720
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
         TabIndex        =   24
         Top             =   30
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   1270
         ButtonWidth     =   1482
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&M.����"
               Key             =   "����"
               Object.ToolTipText     =   "�������ɵĵǼǱ��(Alt+M)"
               Object.Tag             =   "&M.����"
               ImageKey        =   "SendMail"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "�����ɵĵǼǱ�񵼳�ΪExcel�ļ�(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9300
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":1FEE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":220E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":242E
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2BA8
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   9975
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2DC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2FE2
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":3202
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":397C
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   660
      Width           =   2835
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2145
         TabIndex        =   11
         Text            =   "30"
         Top             =   3540
         Width           =   540
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Text            =   "25"
         Top             =   1050
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "&6.���淢��������"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   3240
         Width           =   1845
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2865
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2235
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1635
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.�ȴ�����Ӧ����(��)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3585
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.�˿ں�"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   27
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.�ʼ�������"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.��  ��"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.�û���"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�����˵�ַ"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1425
         Width           =   1080
      End
   End
   Begin VB.Frame fra2 
      Height          =   4320
      Left            =   3210
      TabIndex        =   12
      Top             =   660
      Width           =   7875
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   5
         Left            =   4710
         MaxLength       =   4
         TabIndex        =   30
         Text            =   "500"
         Top             =   525
         Width           =   870
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   1125
         TabIndex        =   17
         Top             =   525
         Width           =   2490
      End
      Begin VB.TextBox txt 
         Height          =   1890
         Index           =   7
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   885
         Width           =   6705
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   4
         Left            =   7425
         Picture         =   "frmScheduleExcel.frx":3B96
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   13
         Left            =   1125
         TabIndex        =   14
         Top             =   165
         Width           =   6300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1050
         Left            =   1125
         TabIndex        =   21
         Top             =   2835
         Width           =   6735
         _cx             =   11880
         _cy             =   1852
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
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -4635
            X2              =   -2850
            Y1              =   -1695
            Y2              =   -1695
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&R.��������"
         Height          =   180
         Index           =   5
         Left            =   3735
         TabIndex        =   29
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&E.�����ʼ�"
         Height          =   180
         Index           =   10
         Left            =   105
         TabIndex        =   16
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&P.������Ա"
         Height          =   180
         Index           =   9
         Left            =   105
         TabIndex        =   20
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&T.�ʼ�����"
         Height          =   180
         Index           =   8
         Left            =   105
         TabIndex        =   18
         Top             =   930
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&N.��������"
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   13
         Top             =   225
         Width           =   900
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1020
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileMail 
         Caption         =   "���ͱ��(&M)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "�������(&S)"
         Shortcut        =   {F2}
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
Attribute VB_Name = "frmScheduleExcel"
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
Private mblnMaining As Boolean

Private Enum mCol
    
    ���� = 66
    �Ա�
    ����
    ��������
    ����״��
    ���֤��
    �����
    ������
    ���￨��
    ������λ
    �����ʼ�
    ����
    ѧ��
    ְҵ
    ����
    �����
    
End Enum

Private Enum mVsfCol
    
    ����
    �Ա�
    ����
    ��������
    ����״��
    ���֤��
    �����ʼ�
    ����
    ѧ��
    ְҵ
    ����
    ������
    ���￨��
    �����
    
End Enum

'�������Զ�����̻���************************************************************************************************

Private Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    strFileTemp = strFileTemp & "���ǼǱ�_" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Private Function NewExcelFile(ByRef strExcelFile As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '����:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim objExcel As Object
    Dim ExWorkbook As Object
    Dim ExWorkSheet As Object
    Dim strParam As String
    Dim lngLoop As Long
    Dim varParam As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim strColChr As String
    
    On Error GoTo errHand
    
    If Val(txt(5).Text) <= 0 Then
        ShowSimpleMsg "����ָ��������������������С��1����"
        LocationObj txt(5)
        Exit Function
    End If

    If Val(txt(5).Text) < vsf.Rows - 1 Then
        ShowSimpleMsg "ָ���������������������������"
        LocationObj txt(5)
        Exit Function
    End If
    
    frmWait.OpenWait Me, "�������ǼǱ�"
    frmWait.WaitInfo = "���ڴ���Excel����..."
    
    Set objExcel = CreateObject("Excel.Application")
    Set ExWorkbook = Nothing
    Set ExWorkSheet = Nothing
    Set ExWorkbook = objExcel.Workbooks().Add
    
    Set ExWorkSheet = ExWorkbook.Worksheets("sheet1")
    
    ExWorkSheet.Name = "��Ա����"
        
    ExWorkSheet.Unprotect "Transaction"                         '����
    objExcel.ActiveWindow.DisplayGridlines = False              'ȡ��������
    
    '�����б���
    ExWorkSheet.Columns("A:A").ColumnWidth = 1
    ExWorkSheet.Range("A3").Value = ""
    
    strParam = "����*,10;�Ա�,5;����,5;��������,10;����״��,10;���֤��*,20;�����*,18;������*,10;���￨��*,15;������λ,15;�����ʼ�,15;����,10;ѧ��,10;ְҵ,10;����,10;�����,10"
    lngRows = Val(txt(5).Text) + 3
    
    varParam = Split(strParam, ";")
    lngCols = UBound(varParam)
    For lngLoop = 0 To lngCols
        strColChr = Chr(lngLoop + 66)
        ExWorkSheet.Range(strColChr & "3").Value = Split(varParam(lngLoop), ",")(0)
        ExWorkSheet.Columns(strColChr & ":" & strColChr).ColumnWidth = Val(Split(varParam(lngLoop), ",")(1))
    Next
    ExWorkSheet.Range("B3:" & Chr(lngCols + 66) & "3").Select
    With objExcel.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Font.Bold = True
        .Font.Size = 9
    End With
    
    '�������
    ExWorkSheet.Range("B1:" & Chr(lngCols + 66) & "1").Select
    With objExcel.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 18
    End With
    objExcel.ActiveCell.FormulaR1C1 = "���������Ա�ǼǱ�"
        
    ExWorkSheet.Range("B2:" & Chr(lngCols + 66) & "2").Select
    With objExcel.Selection
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = True
        .Font.Bold = False
        .Font.Size = 9
        .RowHeight = 30
    End With
    objExcel.ActiveCell.FormulaR1C1 = "ע:������*��ʾ�������룬��û�����������ʾ������Ч��" & vbCrLf & "   ����ʱ�������������֤�Ų����ʷ���ϣ��������µĵ�����"
            
    ExWorkSheet.Range("B4:" & Chr(lngCols + 66) & lngRows).Select
    With objExcel.Selection
        .Locked = False
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = False
        .Font.Size = 9
    End With
    
    '�������
    ExWorkSheet.Range("H4").Select
    objExcel.ActiveWindow.FreezePanes = True
            
    frmWait.WaitInfo = "���ڲ��������ѡ����..."
    
    '���������ѡ����
    ExWorkSheet.Range(Chr(mCol.��������) & "4:" & Chr(mCol.��������) & lngRows).Select
    objExcel.Selection.NumberFormatLocal = "yyyy-mm-dd;@"
    With objExcel.Selection.Validation
        
        .Delete
        .Add 4, 1, 1, "1900-01-01", "3000-01-01"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "����������ȷ��Ч�����ڣ���1980-09-21"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
    ExWorkSheet.Range(Chr(mCol.�Ա�) & "4:" & Chr(mCol.�Ա�) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM �Ա�")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "����������б���ѡ���Ա�"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
        
    ExWorkSheet.Range(Chr(mCol.����״��) & "4:" & Chr(mCol.����״��) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM ����״��")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "����������б���ѡ�����״��"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
'    ExWorkSheet.Range(Chr(mCol.����) & "4:" & Chr(mCol.����) & lngRows).Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM ����")
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "��������"
'        .InputMessage = ""
'        .ErrorMessage = "����������б���ѡ������"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
    
    ExWorkSheet.Range(Chr(mCol.ѧ��) & "4:" & Chr(mCol.ѧ��) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM ѧ��")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "����������б���ѡ��ѧ��"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
    ExWorkSheet.Range(Chr(mCol.ְҵ) & "4:" & Chr(mCol.ְҵ) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM ְҵ")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "����������б���ѡ��ְҵ"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
'    ExWorkSheet.Range(Chr(mCol.����) & "4:" & Chr(mCol.����) & lngRows).Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 3, 1, 1, GetExcelList("SELECT ���� FROM ����")
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "��������"
'        .InputMessage = ""
'        .ErrorMessage = "����������б���ѡ�����"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
'
'    ExWorkSheet.Range(Chr(mCol.����) & "4").Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.����) & "4)>30,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.����) & "4)),FALSE,TRUE))"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "��������"
'        .InputMessage = ""
'        .ErrorMessage = "�������ܺ��зǷ��ַ�(')ͬʱ���Ȳ��ܳ���30���ַ���15�����֣�"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
'    objExcel.Selection.NumberFormatLocal = "@"
'    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.����) & "4:" & Chr(mCol.����) & lngRows), 0
    
    ExWorkSheet.Range(Chr(mCol.����) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.����) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.����) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "�������ܺ��зǷ��ַ�(')ͬʱ���Ȳ��ܳ���20���ַ���10�����֣�"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.����) & "4:" & Chr(mCol.����) & lngRows), 0
            
    ExWorkSheet.Range(Chr(mCol.���֤��) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.���֤��) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.���֤��) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "���֤�Ų��ܺ��зǷ��ַ�(')ͬʱ���Ȳ��ܳ���20���ַ���"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.���֤��) & "4:" & Chr(mCol.���֤��) & lngRows), 0

    ExWorkSheet.Range(Chr(mCol.���￨��) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.���￨��) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.���￨��) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "���￨�Ų��ܺ��зǷ��ַ�(')ͬʱ���Ȳ��ܳ���20���ַ���"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.���￨��) & "4:" & Chr(mCol.���￨��) & lngRows), 0
    
    '�ʼ���ַ
    ExWorkSheet.Range(Chr(mCol.�����ʼ�) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.�����ʼ�) & "4)>50,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.�����ʼ�) & "4)),FALSE,IF(ISNUMBER(FIND(""@""," & Chr(mCol.�����ʼ�) & "4)),TRUE,FALSE)))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "��������"
        .InputMessage = ""
        .ErrorMessage = "�����ʼ���ַ���뺬���ַ�(@)ͬʱ���Ȳ��ܳ���50���ַ���25�����֣�"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.�����ʼ�) & "4:" & Chr(mCol.�����ʼ�) & lngRows), 0
        
     '����������
    ExWorkSheet.Range("B3:" & Chr(lngCols + 66) & lngRows).Select
    With objExcel.Selection
        .Borders(5).LineStyle = -4142
        .Borders(6).LineStyle = -4142
        
        .Borders(7).LineStyle = 1
        .Borders(7).Weight = -4138
        .Borders(7).ColorIndex = 48
        
        .Borders(8).LineStyle = 1
        .Borders(8).Weight = -4138
        .Borders(8).ColorIndex = 48
        
        .Borders(9).LineStyle = 1
        .Borders(9).Weight = -4138
        .Borders(9).ColorIndex = 48
        
        .Borders(10).LineStyle = 1
        .Borders(10).Weight = -4138
        .Borders(10).ColorIndex = 48
        
        .Borders(11).LineStyle = 1
        .Borders(11).Weight = 2
        .Borders(11).ColorIndex = 48
        
        .Borders(12).LineStyle = 1
        .Borders(12).Weight = 2
        .Borders(12).ColorIndex = 48
    End With
    
    frmWait.WaitInfo = "������д�ϴ������Ա����..."
    '��д�ϴ���Ա
    For lngLoop = 1 To vsf.Rows - 1
        '����*,10;�Ա�,5;����,5;��������,10;����״��,10;���֤��*,20;������*,10;���￨��*,15;�����ʼ�,15;����,10;ѧ��,10;ְҵ,10;����,10;�����,10
        '����,1080,1,1,1,;�Ա�,600,1,1,1,;����,600,1,1,1,;��������,990,1,1,1,;����״��,900,1,1,1,;���֤��,1800,1,1,1,;�����ʼ�,1800,1,1,1,;����,900,1,1,1,;ѧ��,900,1,1,1,;ְҵ,900,1,1,1,;����,900,1,1,1,;������,0,1,1,1,;���￨��,0,1,1,1,
        '����,1080,1,1,1,;�Ա�,600,1,1,1,;����,600,1,1,1,;��������,990,1,1,1,;����״��,900,1,1,1,;���֤��,1800,1,1,1,;�����ʼ�,1800,1,1,1,;����,900,1,1,1,;ѧ��,900,1,1,1,;ְҵ,900,1,1,1,;����,900,1,1,1,;������,0,1,1,1,;���￨��,0,1,1,1,;�����,0,1,1,1,"
        
        ExWorkSheet.Range(Chr(mCol.����) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.����)
        ExWorkSheet.Range(Chr(mCol.�Ա�) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.�Ա�)
        ExWorkSheet.Range(Chr(mCol.����) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.����)
        ExWorkSheet.Range(Chr(mCol.��������) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.��������)
        ExWorkSheet.Range(Chr(mCol.����״��) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.����״��)
        ExWorkSheet.Range(Chr(mCol.���֤��) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.���֤��)
        ExWorkSheet.Range(Chr(mCol.�����) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.�����)
        ExWorkSheet.Range(Chr(mCol.������) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.������)
        ExWorkSheet.Range(Chr(mCol.���￨��) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.���￨��)
        ExWorkSheet.Range(Chr(mCol.�����ʼ�) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.�����ʼ�)
        ExWorkSheet.Range(Chr(mCol.����) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.����)
        ExWorkSheet.Range(Chr(mCol.ѧ��) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.ѧ��)
        ExWorkSheet.Range(Chr(mCol.ְҵ) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.ְҵ)
        ExWorkSheet.Range(Chr(mCol.����) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.����)
                
    Next
    
    ExWorkSheet.Range(Chr(mCol.�Ա�) & "3:" & Chr(mCol.�����) & "3").Select
    objExcel.Selection.AutoFilter
    
    '����
    ExWorkSheet.Range(Chr(mCol.����) & "4:" & Chr(mCol.����) & "4").Select
    ExWorkSheet.Protect "transaction", , , , , , , , , , , , , , True
    
    objExcel.ActiveWorkbook.Protect "transaction", True, False
    
    If strExcelFile <> "" Then ExWorkbook.SaveAs strExcelFile
    
    'objExcel.Visible = True
    objExcel.Quit
    
    NewExcelFile = True
    
    Set objExcel = Nothing
    
    frmWait.CloseWait
    
    Exit Function
    
errHand:
    objExcel.Quit
    frmWait.CloseWait
    If ErrCenter = 1 Then Resume
End Function

Private Function GetExcelList(ByVal strSQL As String) As String
    
    Dim rs As New ADODB.Recordset
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            GetExcelList = GetExcelList & "," & zlCommFun.NVL(rs.Fields(0).Value)
            rs.MoveNext
        Loop
    End If
    If GetExcelList = "" Then
        
    Else
        GetExcelList = Mid(GetExcelList, 2)
    End If
End Function

Private Function ValidData() As Boolean
    '���
    
    If Not HaveExcel Then
        MsgBox "�밲װ��Excel����ʹ�ñ����ܡ�", vbCritical, gstrSysName
        Exit Function
    End If
    
    If Trim(txt(4).Text) = "" Then
        MsgBox "����ȷ���ʼ���������"
        LocationObj txt(4)
        Exit Function
    End If
    
    If Val(txt(0).Text) = 0 Then
        MsgBox "�����ʼ��˿ںţ�һ��Ϊ25����"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "����ȷ�������˵ĵ����ʼ���ַ��"
        LocationObj txt(1)
        Exit Function
    End If
    
    
    If Trim(txt(2).Text) = "" Then
        MsgBox "����ȷ���û�����"
        LocationObj txt(2)
        Exit Function
    End If
    
    If Trim(txt(8).Text) = "" Then
        MsgBox "����ȷ����������ʼ���ַ��"
        LocationObj txt(8)
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFileMail.Enabled = vData
    mnuFileSaveAs.Enabled = vData
    
    tbrThis.Buttons("����").Enabled = mnuFileMail.Enabled
    tbrThis.Buttons("����").Enabled = mnuFileSaveAs.Enabled
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
        
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����id
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
                    
    gstrSQL = "select B.����id AS ID,B.����,B.�Ա�,B.����״��,TO_CHAR(C.��������,'yyyy-mm-dd') AS ��������,B.�����ʼ�,C.���֤��,C.����,C.����,C.ѧ��,C.ְҵ,C.����,c.������,c.���￨��,c.����� " & _
                "from ���ǼǼ�¼ A,�����Ա���� B,������Ϣ C " & _
                "where C.����id=B.����id AND A.ID=B.�Ǽ�ID AND B.����id>0 AND A.����=(select max(����) from ���ǼǼ�¼ where ��Լ��λid=[1])"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
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
    
    On Error GoTo errHand
        
    strVsf = "����,1080,1,1,1,;�Ա�,600,1,1,1,;����,600,1,1,1,;��������,990,1,1,1,;����״��,900,1,1,1,;���֤��,1800,1,1,1,;�����ʼ�,1800,1,1,1,;����,900,1,1,1,;ѧ��,900,1,1,1,;ְҵ,900,1,1,1,;����,900,1,1,1,;������,0,1,1,1,;���￨��,0,1,1,1,;�����,0,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Call AppendRows(vsf, lnX, lnY)
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
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


Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case Index
    
    Case 4      '������(��ͬ��λ)ѡ����
        lngKey = Val(cmd(Index).Tag)
        
        gstrSQL = GetPublicSQL(SQL.�������ѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowTxtSelect(Me, txt(13), "����,900,0,1;����,1500,0,1;����,900,0,1;��ַ,3000,0,1", Me.Name & "\�������ѡ��", "�����±���ѡ��һ������/��λ��", rsData, rs, 8790, 5100) Then
            
            lngKey = zlCommFun.NVL(rs("ID").Value, 0)
            txt(13).Text = zlCommFun.NVL(rs("����").Value)
            txt(8).Text = zlCommFun.NVL(rs("�����ʼ�").Value)
              
            cmd(Index).Tag = lngKey
            
            Call ReadData(lngKey)
            
            txt(Index).Tag = ""
        End If
        
        LocationObj txt(13)
        
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyM
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyS
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

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()
    
    glngFormW = 9840
    glngFormH = 6150
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    txt(0).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "������", txt(0).Text)
    txt(1).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�����˵�ַ", txt(1).Text)
    txt(2).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�û���", txt(2).Text)
    txt(3).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", txt(3).Text)
    
    txt(4).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�������", txt(4).Text)
    
'    txt(5).Text = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�������", txt(5).Text))
    txt(6).Text = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ȴ����", txt(6).Text))
    
    chk.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�Ƿ񱣴�����", chk.Value))
    
    txt(7).Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�����", txt(7).Text)
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width + 15
        .Top = fra.Top
        .Width = Me.ScaleWidth - .Left
        .Height = fra.Height
    End With
    
    txt(13).Width = fra2.Width - txt(13).Left - 60 - cmd(4).Width - 30
    cmd(4).Left = txt(13).Left + txt(13).Width + 30
    
'    txt(8).Width = fra2.Width - txt(8).Left - 60
    txt(7).Width = fra2.Width - txt(7).Left - 60
    
    With vsf
        .Width = fra2.Width - .Left - 60
        .Height = fra2.Height - .Top - 60
    End With
    
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnMaining Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "������", txt(0).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�����˵�ַ", txt(1).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�û���", txt(2).Text)
    
    If chk.Value = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", txt(3).Text)
    Else
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", "")
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�������", txt(4).Text)
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�������", Val(txt(5).Text))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ȴ����", Val(txt(6).Text))
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�Ƿ񱣴�����", chk.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ʼ�����", txt(7).Text)
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMail_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    Dim strTmpFile As String
    
    '���
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("����").Enabled = False
    tbrThis.Buttons("�˳�").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    strTmpFile = CreateTmpFile("xls")
    Call NewExcelFile(strTmpFile)
    DoEvents
    
'    gstrSQL = objMail.GetOracleMail(txt(8).Text, "���ǼǱ�", txt(1).Text, txt(4).Text, txt(2).Text, txt(3).Text, "<font color=""ff6633"">����html��ʽ�ʼ�����</font>", strTmpFile, Val(txt(0).Text))
'
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    frmWait.OpenWait Me, "���͵����ʼ�"
    frmWait.WaitInfo = "���������ʼ�������..."

    objMail.ResponseInternal = Val(txt(6).Text)

    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then

        '���͵����ʼ�����

        frmWait.WaitInfo = "���ڷ����������ǼǱ�..."
        blnSuccess = objMail.SendHead(txt(8).Text, txt(2).Text, txt(1).Text, "���ǼǱ�", vbMultipartMixed)
        blnSuccess = objMail.SendMessage(txt(7).Text, vbTextPlain)
        blnSuccess = objMail.SendAttach(strTmpFile)
        blnSuccess = objMail.SendOver
'        blnSuccess = objMail.SendOutLookExMail(txt(8).Text, "���ǼǱ�", txt(7).Text, strTmpFile)

    End If

    frmWait.WaitInfo = "���ڹر��ʼ�������..."

    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("����").Enabled = True
    tbrThis.Buttons("�˳�").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
    '���ɹ�������ʾ
    If blnSuccess = False Then ShowSimpleMsg "���͵����ʼ�ʧ�ܣ�"
    
End Sub

Private Function GetReportMessageHtml(ByVal lngKey As Long, ByVal lng����id As Long) As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim lngLoop1 As Long
    Dim lngLoop2 As Long
    Dim lngLoop3 As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    Dim strSQL As String
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xlTitle style='width:536pt'>��챨�浥</td></tr>"
                        
    strSQL = "SELECT A.����,A.���ʱ��,C.����,B.��첡��id,B.����ʱ��,D.��д�� FROM ���ǼǼ�¼ A,�����Ա���� B,������Ϣ C,���˲�����¼ D WHERE D.ID(+)=B.��첡��id AND C.����id=B.����id AND A.ID=B.�Ǽ�id AND A.ID=[1] AND B.����id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng����id)
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 style='font-weight:700'>�����Ա��<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("����")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>������ڣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("���ʱ��")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>��쵥�ţ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("����")) & "</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr><td colspan=4 class=xl39 style='font-weight:700'>һ����Ŀ����</td></tr>"
            
    '�����������Ŀ����
    
    '1.����
    strSQL = "select DISTINCT C.����,C.ID from �����Ŀҽ�� A,�����Ŀ�嵥 B,���ű� C WHERE A.�嵥ID=B.ID and C.ID=B.ִ�п���id AND A.����id=[1] and B.�Ǽ�id=[2]"
    Set rs1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, lngKey)
    If rs1.BOF Then Exit Function
    
    For lngLoop1 = 1 To rs1.RecordCount
        
        '2.�����Ŀ(��д�˱����)
        strSQL = "select C.����,B.����id,D.��д�� " & _
                        "from ( " & _
                             "SELECT * FROM ����ҽ����¼ WHERE ����id=" & lng����id & " AND �Һŵ�=[1] AND ִ�п���id=[2] AND ������Դ=4 AND ҽ��״̬<>4 AND �������='D' AND ���id IS NULL " & _
                             "Union All " & _
                             "SELECT * FROM ����ҽ����¼ WHERE ����id=" & lng����id & " AND �Һŵ�=[1] AND ִ�п���id=[2] AND ������Դ=4 AND ҽ��״̬<>4 AND �������='C' AND ���id>0 " & _
                             ") A, " & _
                             "����ҽ������ B, " & _
                             "������ĿĿ¼ C, " & _
                             "���˲�����¼ D " & _
                        "Where A.ID = B.ҽ��id " & _
                              "AND B.����id>0 " & _
                              "AND C.ID=A.������ĿID " & _
                              "AND D.ID=B.����id "
        
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlCommFun.NVL(rs("����"))), Val(zlCommFun.NVL(rs1("ID"))))
        If rs2.BOF = False Then
                
            txtInfo.Text = txtInfo.Text & "<tr><td colspan=4 class=xl39 style='font-weight:700'>" & lngLoop1 & "��" & zlCommFun.NVL(rs1("����")) & "</td></tr>"
            
            txtInfo.Text = txtInfo.Text & "<tr>"
            
            For lngLoop2 = 1 To rs2.RecordCount
                
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='font-weight:600'>(" & lngLoop2 & ")" & zlCommFun.NVL(rs2("����")) & "</td>"
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='text-align:right'>���ҽ����<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs2("��д��")) & "</td>"
                txtInfo.Text = txtInfo.Text & "</tr>"
                
                txtInfo.Text = txtInfo.Text & _
                            "<tr><td class=xl25>��Ŀ����</td>" & vbCrLf & _
                            "<td class=xl25>�����</td>" & vbCrLf & _
                            "<td class=xl25>�ο���Χ</td>" & vbCrLf & _
                            "<td class=xl25>��ʾ</td></tr>"
                
                '��������Ŀ�����
                strSQL = _
                    "SELECT * FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "����||DECODE(��־,NULL,'',DECODE(SUBSTR(��־,3,100),'����','','�쳣','(+)','ƫ��','��','ƫ��','��')) AS ����, " & _
                               "�ο�, " & _
                               "DECODE(��־,NULL,'',SUBSTR(��־,3,100)) AS ��ʾ, " & _
                               "�������, " & _
                               "Ԫ������� " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "����, " & _
                               "DECODE(SIGN(INSTR(�ο�,'''')),1,SUBSTR(�ο�,1,INSTR(�ο�,'''')-1),'') AS ��־, " & _
                               "DECODE(SIGN(INSTR(�ο�,'''')),1,SUBSTR(�ο�,INSTR(�ο�,'''')+1,1000),'') AS �ο�, " & _
                               "�������, " & _
                               "Ԫ������� " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "��Ŀ, " & _
                               "DECODE(SIGN(INSTR(����,'''')),1,SUBSTR(����,1,INSTR(����,'''')-1),����) AS ����, " & _
                               "DECODE(SIGN(INSTR(����,'''')),1,SUBSTR(����,INSTR(����,'''')+1,1000),'') AS �ο�, " & _
                               "�������, " & _
                               "Ԫ������� "
                strSQL = strSQL & _
                        "FROM ( " & _
                        "SELECT C.������ AS ��Ŀ,DECODE(A.��������,NULL,NULL,A.��������||' '||DECODE(A.������λ,NULL,'',A.������λ)) AS ����,B.�������,A.�ؼ��� AS Ԫ������� FROM ���˲��������� A,���˲������� B,����������Ŀ C " & _
                        "Where A.����ID = B.ID " & _
                              "AND B.������¼ID=[1] " & _
                              "AND C.ID=A.������ID " & _
                        "))) " & _
                        "Union All " & _
                        "SELECT B.�����ı� AS ��Ŀ,A.����,'' AS �ο�,'' AS ��ʾ,B.�������,0 AS Ԫ������� FROM ���˲����ı��� A,���˲������� B " & _
                        "Where A.����ID = B.ID " & _
                                "And B.������¼ID = [1] " & _
                              "AND Ԫ������ IN (0,-5) " & _
                        ") ORDER BY �������,Ԫ�������"
                        
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("����id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28>" & zlCommFun.NVL(rs3("��Ŀ")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("����")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("�ο�")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("��ʾ")) & "</td></tr>"
                        rs3.MoveNext
                    Next
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28 style='mso-height-source:userset;height:15.0pt'></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td></tr>"
                End If
                                        
                strTmp1 = ""
                strTmp2 = ""
                
                strSQL = "SELECT * FROM �����Ա���� WHERE ����id in (select id from ���˲������� where ������¼id=[1]) ORDER BY ��¼����,��¼���"
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("����id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        
                        If zlCommFun.NVL(rs3("��¼����"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("��������")) & vbCrLf
                        If zlCommFun.NVL(rs3("��¼����"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("�ο�����"))
                        
                        rs3.MoveNext
                    Next
                End If
                
                txtInfo.Text = txtInfo.Text & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>���ۣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>"
                    
                txtInfo.Text = txtInfo.Text & vbCrLf & "<tr><td class=xl39 style='mso-height-source:userset;height:15.0pt'></td></tr>"
                
                rs2.MoveNext
            Next
        End If
        
        rs1.MoveNext
    Next
        
    '�ܼ�
    strTmp1 = ""
    strTmp2 = ""
    
    strSQL = "SELECT * FROM �����Ա���� WHERE ����id in (select id from ���˲������� where ������¼id=[1]) ORDER BY ��¼����,��¼���"
    Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs("��첡��id"))))
    If rs3.BOF = False Then
        For lngLoop3 = 1 To rs3.RecordCount
            
            If zlCommFun.NVL(rs3("��¼����"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("��������")) & vbCrLf
            If zlCommFun.NVL(rs3("��¼����"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("�ο�����"))
            
            rs3.MoveNext
        Next
    End If
            
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=2 class=xl39 style='font-weight:700'>�����ܼ챨��</td>" & vbCrLf & _
        "<td colspan=2 class=xl39 style='text-align:right'>�ܼ�ҽ����<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("��д��")) & "</td></tr>"
        
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���ۣ�<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>���飺<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd") & "</td></tr>"
                
                
    '���
    txtInfo.Text = txtInfo.Text & vbCrLf & "</tr></table></BODY></HTML>"
End Function


Private Sub mnuFileSaveAs_Click()
    Dim strFile As String
       
    If Not HaveExcel Then
        MsgBox "�밲װ��Excel����ʹ�ñ����ܡ�", vbCritical, gstrSysName
        Exit Sub
    End If
    
    dlg.CancelError = True
    
    On Error GoTo ErrHandler
    
    EditChanged = False
    
    dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
    dlg.Filter = "�������(*.xls)| *.xls"
    dlg.FilterIndex = 0
    
    dlg.DialogTitle = "��������ռ�"
    dlg.FileName = App.Path & "\��������ռ�.xls"
    dlg.ShowSave
    If dlg.FileName <> "" Then Call NewExcelFile(dlg.FileName)
            
    EditChanged = True
    
    Exit Sub
    
ErrHandler:
    EditChanged = True
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
    Case "����"
        Call mnuFileSaveAs_Click
    Case "����"
        Call mnuFileMail_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 2 Then txt(2).Tag = "Changed"
    
    If Index = 13 Then
            
        If txt(Index).Tag = "" Then
            Call ResetVsf(vsf)
            txt(Index).Tag = "Changed"
        End If
            
        cmd(4).Tag = ""
        
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index <> 7 Then zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        '����������������а���Enter,��Ҫ������ʷ����
        If txt(Index).Tag = "Changed" And Index = 13 Then
            
            If InStr(txt(Index).Text, "'") Then
                ShowSimpleMsg "�������������зǷ��ַ� ' ��"
                Exit Sub
            End If
            
            gstrSQL = GetPublicSQL(SQL.�������ѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
            
            If ShowTxtFilter(Me, txt(Index), "����,1800,0,0;����,900,0,0;����,900,0,0;��ϵ��,900,0,0;�绰,1200,0,0", Me.Name & "\�������ѡ��", "�������ѡ��һ�����嵥λ", rsData, rs, , , , False) Then
                
                txt(Index).Text = zlCommFun.NVL(rs("����"))
                txt(8).Text = zlCommFun.NVL(rs("�����ʼ�"))
                cmd(4).Tag = zlCommFun.NVL(rs("ID"))
                                                
                Call ReadData(zlCommFun.NVL(rs("ID")))
            Else
                cmd(4).Tag = ""
            End If
            
            txt(Index).Tag = ""

        End If
        
        zlCommFun.PressKey vbKeyTab
        
        If Index = 13 Then
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

