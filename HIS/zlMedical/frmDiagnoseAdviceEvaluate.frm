VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDiagnoseAdviceEvaluate 
   Caption         =   "����������"
   ClientHeight    =   6270
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10500
   Icon            =   "frmDiagnoseAdviceEvaluate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10500
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5040
      Left            =   135
      ScaleHeight     =   5040
      ScaleWidth      =   9420
      TabIndex        =   17
      Top             =   795
      Width           =   9420
      Begin VB.CheckBox chk 
         Caption         =   "��������(&G)"
         Height          =   225
         Left            =   5835
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1305
      End
      Begin VB.Frame fra 
         Height          =   4560
         Left            =   5850
         TabIndex        =   18
         Top             =   405
         Width           =   3450
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   4
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   570
            Width           =   2190
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   4
            Left            =   855
            TabIndex        =   24
            Top             =   915
            Width           =   480
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   5
            Left            =   1575
            TabIndex        =   23
            Top             =   915
            Width           =   480
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   2115
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   915
            Width           =   915
         End
         Begin VB.CommandButton cmd 
            Caption         =   "�Ƴ�����(&R) >>"
            Height          =   350
            Index           =   1
            Left            =   75
            TabIndex        =   13
            Top             =   3090
            Width           =   1440
         End
         Begin VB.CommandButton cmdOpen 
            Height          =   300
            Left            =   3060
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1260
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   855
            TabIndex        =   11
            Top             =   2340
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<< ��ӹ���(&A)"
            Height          =   350
            Index           =   0
            Left            =   75
            TabIndex        =   12
            Top             =   2685
            Width           =   1440
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   855
            TabIndex        =   5
            Top             =   1260
            Width           =   2190
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1620
            Width           =   2190
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   10
            Top             =   1980
            Width           =   2190
         End
         Begin VB.TextBox txt 
            BackColor       =   &H8000000A&
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   3
            Top             =   225
            Width           =   2190
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&2.��  ��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   75
            TabIndex        =   28
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&3.��  ��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   75
            TabIndex        =   27
            Top             =   975
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   1380
            TabIndex        =   26
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lbl 
            Caption         =   "��������ʱ��ֻ������һ��������������Ϊ������������"
            Height          =   435
            Index           =   6
            Left            =   540
            TabIndex        =   21
            Top             =   3615
            Width           =   2745
            WordWrap        =   -1  'True
         End
         Begin VB.Image img 
            Height          =   240
            Left            =   180
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":0CF4
            Top             =   3585
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��λ:s"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   1590
            TabIndex        =   20
            Top             =   3165
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:��ֵ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   1575
            TabIndex        =   19
            Top             =   2805
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&4.��  Ŀ"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   1335
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&5.��  ��"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   7
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&6.��Ŀֵ"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   9
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "&1.��  ��"
            ForeColor       =   &H8000000C&
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   2
            Top             =   300
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   4185
         Left            =   30
         TabIndex        =   0
         Top             =   135
         Width           =   5790
         _cx             =   10213
         _cy             =   7382
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
         Begin VB.Line lnX 
            Index           =   1
            Visible         =   0   'False
            X1              =   1140
            X2              =   2925
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line lnY 
            Index           =   1
            Visible         =   0   'False
            X1              =   1965
            X2              =   1965
            Y1              =   1590
            Y2              =   2805
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   5910
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":127E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13441
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
      Left            =   8115
      Top             =   780
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
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1D2C
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2386
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7515
      Top             =   780
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
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":25A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":27C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":29DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2D2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdviceEvaluate.frx":2F4C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10500
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
         TabIndex        =   16
         Top             =   30
         Width           =   10380
         _ExtentX        =   18309
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
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+S)"
               Object.Tag             =   "&S.����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+R)"
               Object.Tag             =   "&R.����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "����(&R)"
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
Attribute VB_Name = "frmDiagnoseAdviceEvaluate"
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
Private Type Items
    ��Ŀ As String
End Type

Private usrSaveItem As Items

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True

    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False

    End If

    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("����").Enabled = mnuFileRestore.Enabled
        
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next

    Call ResetVsf(vsf)
    Call AppendRows(vsf, lnX, lnY)
    
    On Error GoTo 0
    
    Call InitData
    
    EditChanged = True
    
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, _
                            ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    Dim lngLoop As Long
    
    mblnStartUp = True
    mblnOK = False
                    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
    EditChanged = False
    
    stbThis.Panels(2).Text = "��д�����ϵ���������"
                
    Me.Show 1, frmMain
        
    ShowEdit = mblnOK
    
End Function

Private Function InitSysFlag() As Boolean

    cbo(1).Clear
    cbo(2).Clear
    
    Select Case Mid(lbl(4).Caption, 4)
    Case "������"
        cbo(1).AddItem "[���ֵ]"
        cbo(1).AddItem "[���ֵ]"
        cbo(2).AddItem "[���ֵ]"
        cbo(2).AddItem "[���ֵ]"
    End Select

    Select Case cbo(0).Text
    Case "����"
        cbo(1).AddItem "[ƫ��]"
        cbo(1).AddItem "[ƫ��]"
        cbo(1).AddItem "[�쳣]"
        cbo(2).AddItem "[ƫ��]"
        cbo(2).AddItem "[ƫ��]"
        cbo(2).AddItem "[�쳣]"
    End Select
        
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT B.ID,A.������ AS ����,B.������ AS ��Ŀ,A.��ϵʽ AS ����,A.����ֵ AS ��Ŀֵ,A.�Ա�,A.��ʼ����,A.�������� from ���������� A,����������Ŀ B WHERE A.��Ŀid=B.ID AND A.������=[1] ORDER BY ������"
           
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        
        If zlCommFun.NVL(rs("����").Value) <> "" Then chk.Value = 1
                
        Call LoadGrid(vsf, rs)
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
        
    On Error GoTo errHand
    
    Dim strVsf As String
   
    strVsf = "����,900,1,1,1,;�Ա�,600,1,1,1,;��ʼ����,900,1,1,1,;��������,900,1,1,1,;��Ŀ,2100,1,1,1,;����,900,1,1,1,;��Ŀֵ,900,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)

    lbl(4).Caption = ""
    lbl(5).Caption = ""
'    vsf.ColHidden(0) = True
    vsf.MergeCol(0) = True
    
    With cbo(4)
        .Clear
        .AddItem ""
        .AddItem "1-��"
        .AddItem "2-Ů"
        .ListIndex = 0
    End With
    
    With cbo(3)
        .Clear
        .AddItem "1-��"
        .AddItem "2-��"
        .AddItem "3-��"
        .ListIndex = 0
    End With
    
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

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_����������_DELETE(" & mlngKey & ")"
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            strSQL(ReDimArray(strSQL)) = "ZL_����������_INSERT(" & mlngKey & ",'" & _
                                        vsf.TextMatrix(lngLoop, 0) & "'," & _
                                        Val(vsf.RowData(lngLoop)) & ",'" & _
                                        vsf.TextMatrix(lngLoop, 5) & "','" & _
                                        vsf.TextMatrix(lngLoop, 6) & "','" & _
                                        vsf.TextMatrix(lngLoop, 1) & "','" & _
                                        vsf.TextMatrix(lngLoop, 2) & "','" & _
                                        vsf.TextMatrix(lngLoop, 3) & "')"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub FillOperate(ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strText As String
    
    strText = cbo(0).Text
    
    cbo(0).Clear
    cbo(1).Clear
    Select Case bytMode
    Case 0  '������
        cbo(0).AddItem "����"
        cbo(0).AddItem "����"
        cbo(0).AddItem "С��"
        cbo(0).AddItem "���ڵ���"
        cbo(0).AddItem "С�ڵ���"
        cbo(0).AddItem "������"
        cbo(0).AddItem "�ڷ�Χ��"
    Case 1, 2 '������
        cbo(0).AddItem "����"
        cbo(0).AddItem "����"
        cbo(0).AddItem "С��"
        cbo(0).AddItem "���ڵ���"
        cbo(0).AddItem "С�ڵ���"
        cbo(0).AddItem "������"
        cbo(0).AddItem "����"
    Case 3  '������(�޼���)
        cbo(0).AddItem "����"
        cbo(0).AddItem "������"
        cbo(0).AddItem "����"
'
'        cbo(1).AddItem "����"
'        cbo(1).AddItem "����"
'        cbo(1).ListIndex = 0
    End Select
    
    On Error Resume Next
    
    cbo(0).Text = strText
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
End Sub

Private Sub cbo_Click(Index As Integer)
    Select Case Index
    Case 0
        cbo(2).Visible = (cbo(Index).List(cbo(Index).ListIndex) = "�ڷ�Χ��")
        
        Call InitSysFlag
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Index > 0 Then
            
            If Chr(KeyAscii) = "'" Then KeyAscii = 0
            
            Select Case Val(cmdOpen.Tag)
            Case 0
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.")
            Case 1
                
            Case 2
                
            End Select
        End If
    End If
    
End Sub

Private Sub chk_Click()
    
    If chk.Value = 1 Then
        txt(1).Enabled = True
        txt(1).BackColor = &H80000005
        lbl(3).ForeColor = &H80000012
        
        ResetVsf vsf
        Call AppendRows(vsf, lnX, lnY)
'        vsf.ColHidden(0) = False
    Else
        txt(1).Enabled = False
        txt(1).BackColor = &H8000000A
        lbl(3).ForeColor = &H8000000C
        
        ResetVsf vsf
        Call AppendRows(vsf, lnX, lnY)
'        vsf.ColHidden(0) = True
    End If
    
    EditChanged = True
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim intRow As Long
    
    Select Case Index
    Case 0
        If Val(vsf.RowData(vsf.Rows - 1)) = 0 Then
            intRow = vsf.Rows - 1
        Else
            vsf.Rows = vsf.Rows + 1
            intRow = vsf.Rows - 1
        End If
        
        If chk.Value = 1 Then
            If Trim(txt(1).Text) = "" Then
                ShowSimpleMsg "��������������������"
                Exit Sub
            End If
        End If
        
        If Val(fra.Tag) = 0 Then Exit Sub
        
        vsf.RowData(intRow) = fra.Tag
        vsf.TextMatrix(intRow, 0) = txt(1).Text
        
        vsf.TextMatrix(intRow, 1) = zlCommFun.GetNeedName(cbo(4).Text)
        
        If Trim(txt(4).Text) <> "" Then
            vsf.TextMatrix(intRow, 2) = Trim(txt(4).Text) & zlCommFun.GetNeedName(cbo(3).Text)
        Else
            vsf.TextMatrix(intRow, 2) = ""
        End If
        
        If Trim(txt(5).Text) <> "" Then
            vsf.TextMatrix(intRow, 3) = Trim(txt(5).Text) & zlCommFun.GetNeedName(cbo(3).Text)
        Else
            vsf.TextMatrix(intRow, 3) = ""
        End If
        
        vsf.TextMatrix(intRow, 4) = txt(0).Text
        vsf.TextMatrix(intRow, 5) = cbo(0).Text
        
        Select Case cbo(1).Text
        Case "[���ֵ]", "[���ֵ]", "[ƫ��]", "[ƫ��]", "[�쳣]"
            
        Case Else
            If Val(cmdOpen.Tag) = 0 Then
                cbo(1).Text = Val(cbo(1).Text)
            End If
        End Select
        
        Select Case cbo(2).Text
        Case "[���ֵ]", "[���ֵ]", "[ƫ��]", "[ƫ��]", "[�쳣]"
            
        Case Else
            If Val(cmdOpen.Tag) = 0 Then
                cbo(2).Text = Val(cbo(2).Text)
            End If
        End Select
        
        If cbo(2).Visible Then
            vsf.TextMatrix(intRow, 6) = cbo(1).Text & " �� " & cbo(2).Text
        Else
            vsf.TextMatrix(intRow, 6) = cbo(1).Text
        End If
        
        
        vsf.Col = 0
        vsf.Sort = flexSortGenericAscending
        
        EditChanged = True
        
        Call AppendRows(vsf, lnX, lnY)
        
        LocationObj txt(0)
        
    Case 1
        
        If vsf.Rows <> 2 Then
            vsf.RemoveItem vsf.Row
        Else
            Call ResetVsf(vsf)
        End If
        Call AppendRows(vsf, lnX, lnY)
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
        EditChanged = True
        
    End Select
End Sub

Private Sub cmdOpen_Click()
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
        
    gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(0), "����,1200,0,1;����,1800,0,0;�ٴ�����,1800,0,0", Me.Name & "\������Ŀѡ��", "��ѡ��һ��������Ŀ��", rsData, rs, 8790, 5100) Then
        
        txt(0).Text = zlCommFun.NVL(rs("����").Value)
        fra.Tag = zlCommFun.NVL(rs("ID").Value)
        cmdOpen.Tag = zlCommFun.NVL(rs("����").Value, 0)
        txt(0).Tag = ""
        
        Select Case Val(cmdOpen.Tag)
        Case 0
            lbl(4).Caption = "����:������"
        Case 1
            lbl(4).Caption = "����:������"
        Case 2
            lbl(4).Caption = "����:�޼���"
        Case Else
            lbl(4).Caption = "����:"
        End Select
        
        lbl(5).Caption = "��λ:" & zlCommFun.NVL(rs("��λ").Value)
        
        usrSaveItem.��Ŀ = txt(0).Text
                                
        Call FillOperate(Val(cmdOpen.Tag))
        
        Call InitSysFlag
        
        cbo(1).Text = ""
        cbo(2).Text = ""
        
    End If

    txt(0).SetFocus
End Sub

'���������弰��ؼ����¼�����******************************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyR
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
    
    With picBack
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        
    End With
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = picBack.Width - .Left - fra.Width - 30
        .Height = picBack.Height - .Top
    End With
    
    With chk
        .Left = vsf.Left + vsf.Width + 30
        .Top = 30
    End With
    
    With fra
        .Left = chk.Left
        .Top = chk.Top + chk.Height + 30 - 90
        .Height = picBack.Height - .Top
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
        
    If MsgBox("ȷʵҪ�ָ���ǰ��ѡ��Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData(mlngKey)
    
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    Dim lngKey As Long
            
    If ValidEdit = False Then Exit Sub
    If SaveEdit(lngKey) Then
        mblnOK = True
        EditChanged = False
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

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuFileRestore_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If Index = 0 Then
        txt(Index).Tag = "Changed"
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)

    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt(Index)
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsData As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 0 And txt(Index).Tag <> "" Then
            
            strText = UCase(txt(Index).Text) & "%"
            
            gstrSQL = GetPublicSQL(SQL.������Ŀ����ѡ��)
            
            If ParamInfo.��Ŀ����ƥ�䷽ʽ = 0 Then strTmp = "%" & strText
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
            
            If ShowTxtFilter(Me, txt(Index), "����,900,0,1;����,2400,0,0;Ӣ����,1200,0,0;�ٴ�����,900,0,0", Me.Name & "\������Ŀ����ѡ��", "����±���ѡ��һ����Ŀ", rsData, rs) Then
                
                txt(0).Text = zlCommFun.NVL(rs("����").Value)
                fra.Tag = zlCommFun.NVL(rs("ID").Value)
                cmdOpen.Tag = zlCommFun.NVL(rs("����").Value, 0)
                txt(0).Tag = ""
                usrSaveItem.��Ŀ = txt(0).Text
                
                Call FillOperate(Val(cmdOpen.Tag))
                
                cbo(1).Text = ""
                cbo(2).Text = ""
                
                Select Case Val(cmdOpen.Tag)
                Case 0
                    lbl(4).Caption = "����:������"
                Case 1
                    lbl(4).Caption = "����:������"
                Case 2
                    lbl(4).Caption = "����:�޼���"
                Case Else
                    lbl(4).Caption = "����:"
                End Select
                lbl(5).Caption = "��λ:" & zlCommFun.NVL(rs("��λ").Value)
                
                Call InitSysFlag
                
            Else
                txt(0).Text = usrSaveItem.��Ŀ
                Exit Sub
            End If
        End If
                                
        zlCommFun.PressKey vbKeyTab
        If Index = 0 Then zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 0
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            
            If Chr(KeyAscii) = "*" Then
                KeyAscii = 0
                Call cmdOpen_Click
            End If
            
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If (txt(Index).Tag = "Changed") And Index = 0 Then
        txt(Index).Text = usrSaveItem.��Ŀ
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
    Call SelectRow(vsf, 1, vsf.Row)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

