VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8565
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6615
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   1875
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1845
      End
   End
   Begin VB.PictureBox picSignCheck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   3540
      ScaleHeight     =   2835
      ScaleWidth      =   4725
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   4755
      Begin VB.CommandButton cmdSignAll 
         Caption         =   "ȫ��"
         Height          =   350
         Left            =   270
         TabIndex        =   18
         ToolTipText     =   "ȷ��"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdSignCur 
         Caption         =   "��֤"
         Height          =   350
         Left            =   2790
         TabIndex        =   16
         ToolTipText     =   "ȷ��"
         Top             =   2370
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��"
         Height          =   350
         Left            =   3690
         TabIndex        =   17
         ToolTipText     =   "ȡ��"
         Top             =   2370
         Width           =   840
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSignData 
         Height          =   1635
         Left            =   -30
         TabIndex        =   15
         Top             =   630
         Width           =   4755
         _cx             =   8387
         _cy             =   2884
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendEditor.ctx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ǩ����ʷ��¼����ѡ������֤��Ҳ�ɽ���ȫ����֤��"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   810
         TabIndex        =   14
         Top             =   150
         Width           =   3720
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   120
         Picture         =   "usrTendEditor.ctx":0062
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.PictureBox pic����ȼ� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1965
      Begin VB.ComboBox cbo����ȼ� 
         Height          =   300
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lbl����ȼ� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ģ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.PictureBox picNothing 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1620
      ScaleHeight     =   405
      ScaleWidth      =   1725
      TabIndex        =   9
      Top             =   60
      Width           =   1725
      Begin VB.Label lblNothing 
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ���ˣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   10
         Top             =   90
         Width           =   1875
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   8385
      TabIndex        =   2
      Top             =   510
      Width           =   8385
      Begin MSComctlLib.ListView lvwMultiSel 
         Height          =   1725
         Left            =   2310
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   3043
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   945
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   945
         Begin VB.CommandButton cmdδ��˵�� 
            Caption         =   "�E"
            Height          =   225
            Left            =   630
            TabIndex        =   4
            Top             =   30
            Width           =   255
         End
         Begin VB.ComboBox cbo��λ 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txt���� 
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Vsf 
         Height          =   3975
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8385
         _cx             =   14790
         _cy             =   7011
         Appearance      =   0
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   600
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendEditor.ctx":0CA4
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         Begin VB.PictureBox picSign 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3600
            ScaleHeight     =   195
            ScaleWidth      =   945
            TabIndex        =   19
            Tag             =   "225"
            Top             =   390
            Visible         =   0   'False
            Width           =   975
            Begin VB.Label lbl��֤ǩ�� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��֤ǩ��"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   210
               TabIndex        =   20
               Top             =   0
               Width           =   720
            End
            Begin VB.Image imgSign 
               Height          =   240
               Left            =   -30
               Picture         =   "usrTendEditor.ctx":0D06
               Tag             =   "240"
               Top             =   -30
               Width           =   240
            End
         End
      End
   End
   Begin VB.TextBox txt��ʾ���� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   5970
      MaxLength       =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   90
      Width           =   645
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   4020
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendEditor.ctx":7558
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "usrTendEditor.ctx":DDBA
      Left            =   690
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "usrTendEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public mblnEditable As Boolean

Private objESign As Object
Private mfrmParent As Object
Private mblnInit As Boolean
Private mstrSel As String                   '������:1;����ĳ��Ԫ��:1.1
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mintPreDays As Long
Private mstrMaxDate As String
Private mstrSelItems As String              '�����û��������ӵ��У�����ˢ�º���������
Private mblnCheckVersion As Boolean

Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mbyt����ȼ� As Byte
Private mintӤ�� As Integer
Private mbln���� As Boolean                 '�Ƿ���Ҫ¼������
Private mstrPrivs As String

Private mlngOper As Long                    '�����к�
Private mlngSigner As Long                  'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngGroup As Long                   '���
Private mlngCert As Long                    '֤��ID
Private mlngCertLevel As Long               '��ʿ/��ʿ��ǩ��
Public mstrPigeonhole As String             '�鵵��

Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥

Private Enum ѡ��
    ��
    ��
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Event AfterDataChanged()
Public Event AfterArchiveChanged()
Public Event AfterRefresh()
Public Event AfterSelChange(ByVal lngCert As Long, ByVal strCertLevel As String)
Public Event DbClick(ByVal strData As String)
Public Event AfterRowColChange(ByVal strInfo As String)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'��¼�ϴ�ѡ����,����,�Ա�ˢ�º����¶�λ
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long
Private mbytFontSize As Byte '�����С 9��12

Public Sub ReSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Object
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar
    Dim lngCol As Long, lngReDraw As Long
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))

    UserControl.FontSize = mbytFontSize
    UserControl.FontName = "����"
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = mbytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    Set CtlFont = dkpMain.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = mbytFontSize
    Set dkpMain.PaintManager.CaptionFont = CtlFont
    
        '��ʾ����������
    '------------------------------------------------------------------------------------------------------------------
    lbl����ȼ�.FontSize = mbytFontSize
    cbo����ȼ�.FontSize = mbytFontSize
    cbo����ȼ�.Left = lbl����ȼ�.Left + lbl����ȼ�.Width + 30
    lbl����ȼ�.Top = cbo����ȼ�.Top + (cbo����ȼ�.Height - lbl����ȼ�.Height) \ 2
    cbo����ȼ�.Width = 1575 + IIf(mbytFontSize = 12, 360, 0)
    pic����ȼ�.Width = cbo����ȼ�.Width + cbo����ȼ�.Left
    pic����ȼ�.Height = cbo����ȼ�.Top * 2 + cbo����ȼ�.Height
    txt��ʾ����.FontSize = mbytFontSize
    cbo����.FontSize = mbytFontSize
    picPati.Width = cbo����.Width + cbo����.Left
    picPati.Height = cbo����.Height + cbo����.Top
    
    If Not cbsThis Is Nothing Then
        Set objExtendedBar = cbsThis.Add("�鿴", xtpBarTop)
        objExtendedBar.ContextMenuPresent = False
        objExtendedBar.ShowTextBelowIcons = False
        objExtendedBar.EnableDocking xtpFlagHideWrap
        With objExtendedBar.Controls
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Visible = mblnEditable
            pic����ȼ�.Visible = mblnEditable
            cbrCustom.Handle = pic����ȼ�.hWnd
            cbrCustom.ToolTipText = "����ȼ�"
            
            Set cbrControl = .Add(xtpControlLabel, 0, "��ʾ����")
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Handle = txt��ʾ����.hWnd
            cbrCustom.ToolTipText = "��ʾ�������ڵ�����"
            Set cbrControl = .Add(xtpControlLabel, 0, "����")
            If Not mblnEditable Then cbrControl.Visible = False
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Visible = mblnEditable
            picPati.Visible = mblnEditable
            cbrCustom.Handle = picPati.hWnd
            cbrCustom.ToolTipText = "�����б�"
        End With
        
        For Each objCtrl In cbsThis.Item(cbsThis.Count - 1).Controls
            objCtrl.Delete
        Next
        If Not cbsThis.Item(cbsThis.Count - 1) Is Nothing Then cbsThis.Item(cbsThis.Count - 1).Delete
        cbsThis.Item(2).Visible = mblnEditable
    End If
    
    '��ʼ���б����
    With vsf
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        .FontSize = mbytFontSize
        .FontName = "����"
        .RowHeightMin = BlowUp(IIf(mblnEditable, 600, 300))
        .RowHeightMax = BlowUp(2000)
        .ColWidth(0) = BlowUp(300)
        .ColWidth(1) = BlowUp(1000)
        .ColWidth(2) = BlowUp(800)
        For lngCol = 3 To .Cols - 1
            .ColWidth(lngCol) = BlowUp(900)
        Next lngCol
        Call vsf.AutoSize(0, vsf.Cols - 1)
        .Redraw = lngReDraw
        .Refresh
    End With
    
    cbsThis.RecalcLayout
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If mbytFontSize <> 12 Then Exit Function
    BlowUp = CInt(dblChange + (dblChange * 1 / 3))
End Function

Public Function GetCopyData() As String
    Dim intCol As Integer
    Dim lngOrder As Long
    Dim blnCopy As Boolean, blnDo As Boolean
    Dim strOrder As String, strData As String
    On Error GoTo errHand
    'ֻ������Ч��Ŀ����(������)
    
    If vsf.Row <> vsf.RowSel Then
        MsgBox "��֧�ֶ��и��ƣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    For intCol = vsf.Col To vsf.ColSel
        mrsSelItems.Filter = "��=" & intCol
        If mrsSelItems.RecordCount <> 0 Then
            lngOrder = mrsSelItems!��Ŀ���
            blnCopy = True
        ElseIf vsf.Col = mlngOper Then
            lngOrder = mlngOper
            blnCopy = True
        Else
            blnCopy = False
        End If
        
        If blnCopy Then
            strOrder = strOrder & IIf(Not blnDo, "", ",") & lngOrder
            strData = strData & IIf(Not blnDo, "", ",") & vsf.TextMatrix(vsf.Row, intCol)
            blnDo = True
        End If
    Next
    mrsSelItems.Filter = 0
    
    GetCopyData = strOrder & "|" & strData
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Function

Public Function IsPigeonhole() As Boolean
    IsPigeonhole = (mstrPigeonhole <> "")
End Function

Private Sub cbo����_Click()
    If mblnInit = False Then Exit Sub
    If cbo����.Tag = cbo����.ListIndex Then Exit Sub
    
    cbo����.Tag = cbo����.ListIndex
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), mbyt����ȼ�, mstrPrivs, False, mblnEditable)
End Sub

Private Sub cbo��λ_Click()
    If txt����.Enabled = False Or Val(cbo��λ.Tag) = 1 Then txt����.Text = cbo��λ.Text
End Sub

Private Sub cbo����ȼ�_Click()
    If mblnInit = False Then Exit Sub
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), mbyt����ȼ�, mstrPrivs, False, mblnEditable)
End Sub

'----------------------------------------------------------------
'¼����صĿ���˵����
'�̶�/��ʾ¼��������׾��������
'���������ݺ��¼��򵯳���λ��ʽ
'��*��С�����������ַ������¼�
'��Del�������ǰ�е�����
'-----------------
'�и�ʽ˵��:����,ʱ��,(����,����...,)����,(������...,)ǩ����,ǩ��ʱ��
'����,ʱ���ǹ̶���
'�����ģ�嶨����Ŀ,Ȼ����������,Ҳ�ǹ̶���
'���ѯ������,���Ǵ��ڵ�ǰ��������Ŀ,�Զ�����Ŀ��ӵ������
'�����ǩ����,ǩ��ʱ��,��¼ID,���,��¼��
'-----------------
'�������˵��
'RowData:0-δ�޸�;1-�������޸�
'CellData:0-δ�޸�;1-�������޸�
'-----------------
'ֻ�м�¼IDΪ�յ���,������ɾ������;����,ֻ�����������,ʱ���������
'----------------------------------------------------------------


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim blnClear As Boolean     '�����?
    Dim blnData As Boolean      '��������Ϊ��
    Dim strSymbol As String
    Dim strSelItems As String
    Dim strDelItem As String
    Dim lngOrder As Long
    Dim intRow As Integer, intCol As Integer, intRowSel As Integer, intColSel As Integer
    Dim intRow_ As Integer, intCol_ As Integer
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        mstrSel = vsf.Row & "," & vsf.Col & "," & vsf.RowSel & "," & vsf.ColSel
    Case conMenu_Edit_PASTE
        Call PasteData
    Case conMenu_Edit_Clear
        
        '���ν������ݵ����ҳ���,����colData����Ϊ1,Ȼ����ѡ��Ԫ����������
        blnEnable = picInput.Visible
        intRow = vsf.Row
        intCol = vsf.Col
        intRowSel = vsf.RowSel
        intColSel = vsf.ColSel
        
        If vsf.Row > vsf.RowSel Then intRow = vsf.RowSel: intRowSel = vsf.Row
        If vsf.Col > vsf.ColSel Then intCol = vsf.ColSel: intColSel = vsf.Col
        If intColSel >= mlngSigner Then intColSel = mlngSigner - 1
        
        For intRow_ = intRow To intRowSel
            For intCol_ = intCol To intColSel
                If vsf.TextMatrix(intRow_, intCol_) <> "" Then
                    'ֻ�м�¼IDΪ�յ���,������ɾ������;����,ֻ�����������,ʱ���������
                    If Not (Val(vsf.TextMatrix(intRow_, mlngRecord)) <> 0 And intCol_ <= 2) Then
                        blnClear = CheckVersion(intRow_, intCol_)
                        
                        If blnClear Then
                            vsf.Cell(flexcpData, intRow_, intCol_) = 1
                            vsf.Cell(flexcpText, intRow_, intCol_) = ""
                            vsf.RowData(intRow_) = 1
                            mblnChange = True
                        End If
                    End If
                End If
            Next
        Next
        
        '���ڼ�¼IDΪ��,�����������ݵ���Ч��,ɾ����
        intRowSel = vsf.Rows - 1        '���һ����Զ��ɾ,���������հ���,�����û�¼��
        intColSel = mlngSigner - 1
        For intRow = intRowSel To 1 Step -1
            blnData = False
            For intCol = IIf(Val(vsf.RowData(intRow)) = 0, 1, 3) To intColSel
                If vsf.TextMatrix(intRow, intCol) <> "" Then
                    blnData = True
                    Exit For
                End If
            Next
            If Not blnData Then
                If Val(vsf.TextMatrix(intRow, mlngRecord)) <> 0 Then   '��ʷ��������
                    vsf.RowHidden(intRow) = True
                Else
                    If intRow <> vsf.Rows - 1 Then
                        vsf.RemoveItem intRow               '�¼�¼ɾ��
                    End If
                End If
            End If
        Next
        
        mblnShow = False
        picInput.Visible = False
        
        '���ѡ������
        vsf.RowSel = vsf.Row
        vsf.ColSel = vsf.Col
        vsf.SetFocus
        If blnEnable Then Call Vsf_EnterCell
        If mblnChange Then RaiseEvent AfterDataChanged
    Case conMenu_Edit_SPECIALCHAR
        strSymbol = frmInsSymbol.ShowMe(False, 0)
        txt����.Text = txt����.Text & strSymbol
    Case conMenu_Edit_Append
        '��������ǩ����֮�����,������ʱ��ӵ���Ŀ,�ⲿ����Ŀ�ǰ���Ŀ��Ŵ�С˳����ӵ�,���,���ֹ����ʱ,ҲӦ�ñ�֤��˳��,����ˢ�º���˳�����仯
        With mrsSelItems
            '�õ���ѡ����Ŀ������嵥
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                strSelItems = strSelItems & "," & !��Ŀ���
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
        strSelItems = strSelItems & ","
        
        strSelItems = frmTendItemChoose.ShowSelect(strSelItems, cbo����ȼ�.ListIndex, cbo����.ItemData(cbo����.ListIndex), mlng����ID)
        If strSelItems = "" Then Exit Sub
        mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & strSelItems
        
        Call InsertColumn(strSelItems)
    Case conMenu_Edit_Delete
        '�����ѯ�б���������������ɾ��
        intCol = vsf.Col
        intRowSel = vsf.Rows - 1
        For intRow = vsf.Row To intRowSel
            If vsf.TextMatrix(intRow, intCol) <> "" Or vsf.Cell(flexcpData, intRow, intCol) <> 0 Then
                MsgBox "��ǰ��Ŀ�����ݣ�������ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
        
        Call DeleteColumn(intCol)
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ǩ������ݲ��������κδ���
    
    If mblnInit = False Then Exit Sub
    Select Case Control.ID
    Case conMenu_Edit_PASTE
        Control.Visible = mblnEditable
        Control.Enabled = (mstrSel <> "") And Not IsPigeonhole And mblnEditable And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "��ʿ��"
    Case conMenu_Edit_Copy, conMenu_Edit_SPECIALCHAR, conMenu_Edit_Append
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And (InStr(1, mstrPrivs, "�����¼�Ǽ�") <> 0)
    Case conMenu_Edit_Clear 'ǩ�������ݲ��������
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And (InStr(1, mstrPrivs, "�����¼�Ǽ�") <> 0) And mblnCheckVersion And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "��ʿ��"
        'If Control.Enabled Then Control.Enabled = (Vsf.TextMatrix(Vsf.Row, mlngSigner) = "")
        
        '����Ƕ�ѡ,���������
        If vsf.RowSel <> vsf.Row Or vsf.ColSel <> vsf.Col Then Control.Enabled = True
    Case conMenu_Edit_Delete
        Dim blnDel As Boolean
        If mrsSelItems.State = 1 Then
            mrsSelItems.Filter = "��=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                blnDel = (mrsSelItems!�̶� = 0)
            End If
            mrsSelItems.Filter = 0
        End If
        Control.Visible = mblnEditable
        Control.Enabled = Not IsPigeonhole And mblnEditable And blnDel And (InStr(1, mstrPrivs, "�����¼�Ǽ�") <> 0) And vsf.TextMatrix(vsf.Row, mlngCertLevel) <> "��ʿ��"
    End Select
End Sub

Private Sub cmdδ��˵��_Click()
    If cbo��λ.Visible Then
        If Val(cbo��λ.Tag) = 0 Then
            Call txt����_KeyDown(vbKeyDown, vbShiftMask)
        Else
            Call txt����_KeyDown(vbKeyDown, 0)
            txt����.Text = ""
            txt����.SetFocus
        End If
    Else
        Call txt����_KeyDown(vbKeyW, vbCtrlMask)
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Where B.Ӧ�÷�ʽ<>0 " & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitBill()
    Dim intCol As Integer, intCols As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    '��ʼ���ڴ��¼��
    strFields = "��," & adDouble & ",18|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2"
    Call Record_Init(mrsSelItems, strFields)
    strFields = "��|��Ŀ���|��Ŀ����|�̶�"
    
    '�����ģ���趨����Ŀ
    strSQL = " Select B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ,B.��Ŀ����,1 AS �̶�" & _
             " From ������Ŀģ�� A,�����¼��Ŀ B" & _
             " Where a.��Ŀ��� = b.��Ŀ��� And B.Ӧ�÷�ʽ<>0 And A.����ID=[3] And A.����ȼ� = [1] And B.���ò��� IN (0,[2])" & _
             " And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[3])))" & _
             " Order by A.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����ģ���趨����Ŀ", cbo����ȼ�.ListIndex, IIf(cbo����.ItemData(cbo����.ListIndex) = 0, 1, 2), mlng����ID)
    If rsTemp.RecordCount = 0 Then
        '����ǰ�Ĺ�����ȡ��Ŀ�嵥��¼��
        strSQL = " Select B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ,B.��Ŀ����,0 AS �̶�" & _
                 " From �����¼��Ŀ B" & _
                 " Where B.Ӧ�÷�ʽ<>0 And B.����ȼ�>=[1] And B.���ò��� IN (0,[2])" & _
                 " And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[3])))" & _
                 " Order by B.��Ŀ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ǰ�Ĺ�����ȡ��Ŀ�嵥��¼��", cbo����ȼ�.ListIndex, IIf(cbo����.ItemData(cbo����.ListIndex) = 0, 1, 2), mlng����ID)
    End If
    
    With vsf
        intCols = .Cols - 1
        For intCol = 1 To intCols
            .ColHidden(intCol) = False
        Next
        
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = rsTemp.RecordCount + .FixedCols + 3     '��������ʱ����,�ټ��Ϲ̶���������
        .RowHeightMin = IIf(mblnEditable, 600, 300)
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .WordWrap = True
        
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "ʱ��"
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 600
        .ColWidth(2) = 800
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        intCol = 3
        Do While Not rsTemp.EOF
            If rsTemp!��Ŀ���� Like "����ѹ*" And .TextMatrix(0, intCol - 1) Like "����ѹ*" Then
                .TextMatrix(0, intCol - 1) = "Ѫѹ" & IIf(NVL(rsTemp!��Ŀ��λ) = "", "", vbCrLf & "(" & rsTemp!��Ŀ��λ & ")")
                .Cols = .Cols - 1
                intCol = intCol - 1
            Else
                .TextMatrix(0, intCol) = rsTemp!��Ŀ���� & IIf(NVL(rsTemp!��Ŀ��λ) = "", "", vbCrLf & "(" & rsTemp!��Ŀ��λ & ")")
            End If
            .ColWidth(intCol) = 900
            .ColAlignment(intCol) = IIf(rsTemp!��Ŀ���� = 0, flexAlignCenterCenter, flexAlignLeftTop)       '�����������ʾ,���������û�¼���������ʾ
            
            '��Ŀǰ��ѡ�����Ŀ�����ڴ��¼����
            strValues = intCol & "|" & rsTemp!��Ŀ��� & "|" & rsTemp!��Ŀ���� & "|" & rsTemp!�̶�
            Call Record_Add(mrsSelItems, strFields, strValues)
            
            intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        '.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .MergeCells = flexMergeFree
        .WordWrap = True
        
        '��Ŀǰ��ѡ�����Ŀ�����ڴ��¼����
        strValues = .Cols - 1 & "|0|����|1"
        Call Record_Add(mrsSelItems, strFields, strValues)
        
        mlngOper = .Cols - 1
        .TextMatrix(0, .Cols - 1) = "����"
        .TextMatrix(1, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        .TextMatrix(1, 2) = Format(zlDatabase.Currentdate, "HH:mm")
    End With
    
    '����Ƿ���Ҫ¼������
    mrsSelItems.Filter = "��Ŀ���=-1"
    mbln���� = (mrsSelItems.RecordCount <> 0)
    mrsSelItems.Filter = 0
End Sub

Private Sub ReadData()
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    
    Dim int����Ӧ�� As Integer
    Dim strStart As String, strEnd As String
    Dim rsColumns As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ȡ���ڶ����������
    
    mrsItems.Filter = "��Ŀ���=-1"
    If mrsItems.RecordCount <> 0 Then
        int����Ӧ�� = mrsItems!Ӧ�÷�ʽ
    End If
    mrsItems.Filter = 0
    strStart = Format(DateAdd("d", -1 * Val(txt��ʾ����.Text), zlDatabase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
    
    '����Ƿ�鵵
    gstrSQL = " Select �鵵�� From ���˻����¼ Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And Rownum<2"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�鵵", mlng����ID, mlng��ҳID, cbo����.ItemData(cbo����.ListIndex))
    If rsTemp.RecordCount <> 0 Then mstrPigeonhole = NVL(rsTemp!�鵵��)
    
    '1������ȡ����ѯʱ�䷶Χ���Լ���ӵ���Ŀ,���μӵ������
    gstrSQL = " Select Distinct Y.��Ŀ���,Y.��Ŀ���� From (" & _
                    " Select A.��Ŀ��� " & _
                    " From ���˻������� A,���˻����¼ C" & _
                    " Where C.ID = A.��¼id AND A.��¼���� =1 AND C.������Դ = 2 AND ((NVL(A.��¼���,0) <> 1 And a.��Ŀ���>0) or a.��Ŀ���=-1 ) " & _
                         " AND C.����ʱ�� Between [1] And [2] And C.����ID=[3] And C.��ҳID=[4]" & _
                    "       " & _
                    "      ) X,�����¼��Ŀ Y " & _
              " Where Y.��Ŀ��� = X.��Ŀ��� AND nvl(Y.����ȼ�,3) >=[6] And Nvl(y.Ӧ�÷�ʽ,0)=1 And Nvl(y.���ò���,0) In (0,[7]) And (Y.���ÿ���=1 Or (Y.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=Y.��Ŀ��� And D.����id=[5])))  " & _
              " Order By Y.��Ŀ���"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsColumns = zlDatabase.OpenSQLRecord(gstrSQL, "����ȡ����ѯʱ�䷶Χ���Լ���ӵ���Ŀ,���μӵ������", CDate(strStart), CDate(strEnd), _
                mlng����ID, mlng��ҳID, mlng����ID, cbo����ȼ�.ListIndex, IIf(cbo����.ItemData(cbo����.ListIndex) = 0, 1, 2))
    Call AddColumns(rsColumns)
    
    '���û�ѡ�������ӽ�ȥ
    If mstrSelItems <> "" Then
        arrColumn = Split(mstrSelItems, vbCrLf)
        intEnd = UBound(arrColumn)
        For intStart = 0 To intEnd
            Call InsertColumn(arrColumn(intStart))
        Next
    End If
    vsf.Cell(flexcpAlignment, 0, 0, 0, vsf.Cols - 1) = flexAlignCenterCenter
    
    '2����ȡ����
    gstrSQL = " Select X.* From ("
    If int����Ӧ�� = 2 Then
        gstrSQL = gstrSQL & _
                    "Select A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                        "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,NVL(D.��Ŀ����,to_char(D.�޸�ʱ��,'yyyy-MM-dd hh24:mi:ss')) As ǩ��ʱ��,NVL(D.��¼����,'��ʿ') AS ǩ������," & _
                        "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,b.��¼���� As ���,b.��¼���," & _
                        "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,a.��¼�� " & _
                    " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                    " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                         " AND A.��¼���� =1 AND C.������Դ = 2 AND NVL(A.��¼���,0) <> 1 " & _
                         " And D.��¼����(+)=5 And D.��¼ID(+)=C.ID And D.��ֹ�汾(+) Is NULL" & _
                         " AND C.����ʱ�� Between [1] And [2] And C.����ID=[3] And C.��ҳID=[4] and C.Ӥ��=[8]"
    Else
        gstrSQL = gstrSQL & _
                    "Select A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                        "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,NVL(D.��Ŀ����,to_char(D.�޸�ʱ��,'yyyy-MM-dd hh24:mi:ss')) As ǩ��ʱ��,NVL(D.��¼����,'��ʿ') AS ǩ������," & _
                        "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,Decode(a.��Ŀ���,2,'',-1,'',b.��¼����) As ���,Decode(a.��Ŀ���,2,0,-1,0,b.��¼���) As ��¼���," & _
                        "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,a.��¼�� " & _
                    " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                    " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                         " AND A.��¼���� =1 AND C.������Դ = 2 AND ((NVL(A.��¼���,0) <> 1 And a.��Ŀ���>0) or a.��Ŀ���=-1 or (a.��Ŀ���=0 and a.��¼����=4)) " & _
                         " And D.��¼����(+)=5 And D.��¼ID(+)=C.ID And D.��ֹ�汾(+) Is NULL" & _
                         " AND C.����ʱ�� Between [1] And [2] And C.����ID=[3] And C.��ҳID=[4] and C.Ӥ��=[8]"
    End If
    gstrSQL = gstrSQL & _
                "       And a.��ֹ�汾 Is Null And b.��ֹ�汾 Is Null " & _
                "       And Decode(a.��Ŀ���,2,-1,a.��Ŀ���)=b.��Ŀ���(+)) X,�����¼��Ŀ Y " & _
                "Where Y.��Ŀ��� = X.��Ŀ��� AND nvl(Y.����ȼ�,3) >=[6] And Nvl(y.Ӧ�÷�ʽ,0)=1 And Nvl(y.���ò���,0) In (0,[7]) And (Y.���ÿ���=1 Or (Y.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=Y.��Ŀ��� And D.����id=[5])))  "
    
    '����������Ŀ
    gstrSQL = gstrSQL & _
                " UNION " & _
                " Select A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                    "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,NVL(D.��Ŀ����,to_char(D.�޸�ʱ��,'yyyy-MM-dd hh24:mi:ss')) As ǩ��ʱ��,NVL(D.��¼����,'��ʿ') AS ǩ������," & _
                    "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,Decode(a.��Ŀ���,2,'',-1,'',b.��¼����) As ���,Decode(a.��Ŀ���,2,0,-1,0,b.��¼���) As ��¼���," & _
                    "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,a.��¼�� " & _
                " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                     " AND A.��¼���� =4 AND C.������Դ = 2 And a.��ֹ�汾 Is Null And b.��ֹ�汾 Is Null And D.��ֹ�汾(+) Is NULL" & _
                     " And D.��¼����(+)=5 And D.��¼ID(+)=C.ID" & _
                     " AND C.����ʱ�� Between [1] And [2] And C.����ID=[3] And C.��ҳID=[4] And C.Ӥ��=[8]"
    
    gstrSQL = " Select * From (" & gstrSQL & ") Order By �������,��¼ID,��¼���,DECODE(��Ŀ���,0,999,��Ŀ���)"

    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", CDate(strStart), CDate(strEnd), _
                mlng����ID, mlng��ҳID, mlng����ID, cbo����ȼ�.ListIndex, IIf(cbo����.ItemData(cbo����.ListIndex) = 0, 1, 2), cbo����.ItemData(cbo����.ListIndex))

    '׼���������(����û�е���Ŀ,ֱ���ڱ�������Ӹ���,ͬʱ�����ڲ���¼��
    Call ShowData(rsData)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteColumn(ByVal intCol As Integer)
    Dim lngOrder As Long
    Dim strName As String
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    'ɾ��ָ������
    
    mrsSelItems.Filter = "��=" & intCol
    lngOrder = mrsSelItems!��Ŀ���
    strName = mrsSelItems!��Ŀ����
    mrsSelItems.Filter = 0
    
    'ɾ����
    vsf.ColPosition(intCol) = vsf.Cols - 1
    vsf.Cols = vsf.Cols - 1
    '�����ڲ���¼��
    With mrsSelItems
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !�� > intCol Then
                !�� = !�� - 1
                .Update
            ElseIf !�� = intCol Then
                .Delete
            Else
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    '���ģ������ĸ���
    If mlngOper > intCol Then mlngOper = mlngOper - 1
    mlngSigner = mlngSigner - 1
    mlngSignTime = mlngSignTime - 1
    mlngRecord = mlngRecord - 1
    mlngGroup = mlngGroup - 1
    mlngCert = mlngCert - 1
    mlngCertLevel = mlngCertLevel - 1
    
    arrColumn = Split(mstrSelItems, vbCrLf)
    intEnd = UBound(arrColumn)
    mstrSelItems = ""
    For intStart = 0 To intEnd
        If Val(Split(arrColumn(intStart), "|")(0)) <> lngOrder Then
            mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & arrColumn(intStart)
        End If
    Next
End Sub

Private Sub InsertColumn(ByVal strSelItems As String)
    Dim lngOrder As Long
    
    '����Ѵ��ڸ������˳�
    mrsSelItems.Filter = "��Ŀ���=" & Val(Split(strSelItems, "|")(0))
    If mrsSelItems.RecordCount <> 0 Then
        mrsSelItems.Filter = 0
        Exit Sub
    End If
    
    '���û�ѡ�����Ŀ��ӵ������
    mrsItems.Filter = "��Ŀ���=" & Val(Split(strSelItems, "|")(0))
    vsf.Cols = vsf.Cols + 1
    vsf.TextMatrix(0, vsf.Cols - 1) = Split(strSelItems, "|")(1) & IIf(NVL(mrsItems!��Ŀ��λ) = "", "", vbCrLf & "(" & mrsItems!��Ŀ��λ & ")")
    vsf.ColAlignment(vsf.Cols - 1) = IIf(mrsItems!��Ŀ���� = 0, flexAlignCenterCenter, flexAlignLeftTop)       '�����������ʾ,���������û�¼���������ʾ
    mrsItems.Filter = 0
    'Vsf.Cell(flexcpAlignment, 0, Vsf.Cols - 1, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter  '���н�������
        
    'ȡ���������е���Ŀ���
    With mrsSelItems
        .Filter = "��>" & mlngOper
        .Sort = "��"
        Do While Not .EOF
            If !��Ŀ��� > Val(Split(strSelItems, "|")(0)) Then
                lngOrder = !��
                Exit Do
            End If
            .MoveNext
        Loop
        If lngOrder = 0 Then lngOrder = mlngSigner  'û����,˵��û�������Ŀ,ȡǩ����
        
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    vsf.ColPosition(vsf.Cols - 1) = lngOrder      'ǩ�����п�ʼ������
    '�����ڲ���¼��
    With mrsSelItems
        Do While Not .EOF
            If !�� >= lngOrder Then
                !�� = !�� + 1
                .Update
            End If
            .MoveNext
        Loop
    End With
    strValues = lngOrder & "|" & Split(strSelItems, "|")(0) & "|" & Split(strSelItems, "|")(1) & "|0"
    Call Record_Add(mrsSelItems, strFields, strValues)
    '���ģ������ĸ���
    mlngSigner = mlngSigner + 1
    mlngSignTime = mlngSignTime + 1
    mlngRecord = mlngRecord + 1
    mlngGroup = mlngGroup + 1
    mlngCert = mlngCert + 1
    mlngCertLevel = mlngCertLevel + 1
End Sub

Private Sub AddColumns(ByVal rsColumns As ADODB.Recordset)
    '����ʷ�����д��ڵĶ�������ӵ������
    With rsColumns
        Do While Not .EOF
            mrsSelItems.Filter = "��Ŀ���=" & !��Ŀ���
            If mrsSelItems.RecordCount = 0 Then
                mrsItems.Filter = "��Ŀ���=" & !��Ŀ���
                vsf.Cols = vsf.Cols + 1
                vsf.TextMatrix(0, vsf.Cols - 1) = .Fields("��Ŀ����").Value & IIf(NVL(mrsItems!��Ŀ��λ) = "", "", vbCrLf & "(" & mrsItems!��Ŀ��λ & ")")
                vsf.ColAlignment(vsf.Cols - 1) = IIf(mrsItems.Fields("��Ŀ����").Value = 0, flexAlignCenterCenter, flexAlignLeftTop)
                mrsItems.Filter = 0
                
                strValues = vsf.Cols - 1 & "|" & !��Ŀ��� & "|" & !��Ŀ���� & "|0"
                Call Record_Add(mrsSelItems, strFields, strValues)
            End If
            .MoveNext
        Loop
    End With
    
    '�̶�����ǩ����,ǩ��ʱ����
    With vsf
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "ǩ����"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSigner = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "ǩ��ʱ��"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSignTime = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "֤��ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngCert = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "��¼ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngRecord = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "ǩ������"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngCertLevel = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "���"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngGroup = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "��¼��"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    mrsSelItems.Filter = 0
End Sub

Private Sub ShowData(ByVal rsData As ADODB.Recordset)
    On Error GoTo errHand
    Dim lngRow As Long
    Dim lngRecord As Long   '��¼ID
    Dim lngGroup As Long    '���
    Dim strData As String
    Dim strTime As String
    Dim lng��ֹ�汾 As Long, bln��ɫ As Boolean
    Dim rsTemp As New ADODB.Recordset   '��ȡ��ǰ��¼������ֹ�汾
    
    '��ѭ��д����
    lngRow = 1
    With rsData
        Do While Not .EOF
            If lngRecord <> !��¼ID Or lngGroup <> !��¼��� Then
                '��ȡ��ǰ��¼������ֹ�汾
                gstrSQL = " Select max(��ʼ�汾),Max(��ֹ�汾) From ���˻������� Where ��¼ID=[1]"
                If mblnMoved_HL Then
                    gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
                    gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
                End If
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ��¼������ֹ�汾", CLng(!��¼ID))
                lng��ֹ�汾 = NVL(rsTemp.Fields(0).Value, 1)
                If lng��ֹ�汾 < NVL(rsTemp.Fields(1).Value, 1) Then lng��ֹ�汾 = NVL(rsTemp.Fields(1).Value, 1)
                
                '�µļ�¼
                If lngRecord <> 0 Then
                    '������
                    lngRow = lngRow + 1
                    If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
                    vsf.TextMatrix(lngRow, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                    vsf.TextMatrix(lngRow, 2) = Format(zlDatabase.Currentdate, "HH:mm")
                    'Vsf.Cell(flexcpAlignment, lngRow, 0, lngRow, Vsf.Cols - 1) = flexAlignCenterCenter
                Else
                    '��һ����¼
                End If
                strTime = Format(!�������, "yyyy-MM-dd HH:mm")
                
                '��д��ǩ���˼�ǩ��ʱ��
                lngRecord = !��¼ID
                lngGroup = !��¼���
                bln��ɫ = True
                If Not IsNull(!ǩ����) Then
                    bln��ɫ = False
                    vsf.Cell(flexcpPicture, lngRow, 0) = imgRow.ListImages(1).Picture
                End If
                vsf.Cell(flexcpPictureAlignment, lngRow, 0) = flexAlignCenterCenter
                vsf.TextMatrix(lngRow, 1) = Split(strTime, " ")(0)
                vsf.TextMatrix(lngRow, 2) = Split(strTime, " ")(1)
                vsf.TextMatrix(lngRow, mlngCert) = Val(NVL(.Fields("֤��ID").Value, 0))
                vsf.TextMatrix(lngRow, mlngCertLevel) = NVL(.Fields("ǩ������").Value)
                vsf.TextMatrix(lngRow, mlngSigner) = NVL(.Fields("ǩ����").Value)
                vsf.TextMatrix(lngRow, mlngSignTime) = Format(.Fields("ǩ��ʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                vsf.TextMatrix(lngRow, mlngRecord) = CLng(.Fields("��¼ID").Value)
                vsf.TextMatrix(lngRow, mlngGroup) = CLng(.Fields("��¼���").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 1) = NVL(.Fields("��¼��").Value)
                vsf.RowData(lngRow) = 0
                
                If bln��ɫ Then 'ǩ����Ϊ��,����ֹ�汾����1,��˵����Ҫ��ɫ;�ſ��������������ݲ���Ҫ��ɫ�����
                    bln��ɫ = (lng��ֹ�汾 > 1)
                End If
            End If
            
            '��д����ͨ�Ļ�����Ŀ
            If !��Ŀ��� <> 0 Then
                '���δ��˵����Ϊ��,��ʾδ��˵��
                If Not IsNull(.Fields("δ��˵��").Value) Then
                    strData = .Fields("δ��˵��").Value
                Else
                    strData = NVL(.Fields("��¼���").Value)
                    If Not IsNull(.Fields("���").Value) Then
                        strData = strData & "/" & .Fields("���").Value
                    End If
                    If Not IsNull(.Fields("��λ").Value) Then
                        strData = .Fields("��λ").Value & ":" & strData
                    ElseIf !��Ŀ��� = 1 Then
                        strData = "Ҹ��:" & strData
                    End If
                End If
                
                mrsSelItems.Filter = "��Ŀ���=" & !��Ŀ���
                If mrsSelItems.RecordCount <> 0 Then
                    If !��Ŀ��� = 5 Then   '����ѹ,�����Ӧ��Ԫ��������,��˵������������ѹ,��/�����ʾ
                        If vsf.TextMatrix(lngRow, mrsSelItems!��) <> "" Then
                            vsf.TextMatrix(lngRow, mrsSelItems!��) = vsf.TextMatrix(lngRow, mrsSelItems!��) & "/" & strData
                        Else
                            vsf.TextMatrix(lngRow, mrsSelItems!��) = strData
                        End If
                    Else
                        vsf.TextMatrix(lngRow, mrsSelItems!��) = strData
                    End If
                End If
            Else
                '��д������
                strData = NVL(.Fields("��¼���").Value)
                mrsSelItems.Filter = "��Ŀ���=0"
                If mrsSelItems.RecordCount <> 0 Then
                    vsf.TextMatrix(lngRow, mrsSelItems!��) = strData
                End If
            End If
            
            '��ɫ(��������)
            If !ʵ�ʰ汾 = lng��ֹ�汾 And bln��ɫ Then
                vsf.Cell(flexcpForeColor, lngRow, mrsSelItems!��) = &HFF&
            End If
            
            .MoveNext
        Loop
    End With
    mrsSelItems.Filter = 0
    
    '���ӿհ���
    If Val(vsf.TextMatrix(vsf.Rows - 1, mlngRecord)) <> 0 Then
        lngRow = lngRow + 1
        If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
        vsf.TextMatrix(lngRow, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        vsf.TextMatrix(lngRow, 2) = Format(zlDatabase.Currentdate, "HH:mm")
        'Vsf.Cell(flexcpAlignment, lngRow, 0, lngRow, Vsf.Cols - 1) = flexAlignCenterCenter
    End If
    
    'ʹ��CellData�������޸ı�־
    vsf.Cell(flexcpData, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = 0
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Sub

Private Sub PasteData()
    Dim arrSel, arrData, arrRow
    Dim strSource As String
    Dim intRow As Integer, intCol As Integer
    Dim intSourceRow As Integer, intSourceCol As Integer, intSourceRowSel As Integer, intSourceColSel As Integer
    '����ճ���������Ƿ��뿽�������ݴ����ص����㷨��������
    
    arrSel = Split(mstrSel, ",")
    intSourceRow = arrSel(0)
    intSourceCol = arrSel(1)
    intSourceRowSel = arrSel(2)
    intSourceColSel = arrSel(3)
    '�������ѡ���,����Ҫ����һ����ʼ��,��,��ֹ��,��
    If intSourceRow > intSourceRowSel Then intRow = intSourceRow: intSourceRow = intSourceRowSel: intSourceRowSel = intRow
    If intSourceCol > intSourceColSel Then intCol = intSourceCol: intSourceCol = intSourceColSel: intSourceColSel = intCol
    'ǩ����,ǩ��ʱ��,��¼ID,��������в�����
    'If intSourceColSel > Vsf.Cols - 5 Then intSourceColSel = Vsf.Cols - 5
    If intSourceColSel >= mlngSigner Then intSourceColSel = mlngSigner - 1
    
    '��ճ������ʼ�б����뿽������ʼ����ͬ,������ִ��ճ������
    If vsf.Col <> intSourceCol Then
        MsgBox "��ճ�����б����븴�Ƶ���ʼ����ͬ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If vsf.Row = intSourceRow Then Exit Sub
    
    '�õ���ճ������
    If vsf.Row > intSourceRow And vsf.Row <= intSourceRowSel Then
        If MsgBox("����ѡ���ճ�������븴�������غ��ˣ���ȷ��Ҫ����ճ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '�����������������ʱд�������
    For intRow = intSourceRow To intSourceRowSel
        strSource = strSource & IIf(intRow = intSourceRow, "", "|����|")
        For intCol = intSourceCol To intSourceColSel
            strSource = strSource & IIf(intCol = intSourceCol, "", "|С��|") & vsf.TextMatrix(intRow, intCol)
        Next
    Next
    
    '���и���(�����в�����)
    If strSource = "" Then Exit Sub
    arrData = Split(strSource, "|����|")
    intSourceRowSel = vsf.Row + (intSourceRowSel - intSourceRow)
    For intRow = vsf.Row To intSourceRowSel
        arrRow = Split(arrData(intRow - vsf.Row), "|С��|")
        If intRow > vsf.Rows - 1 Then Exit For
        For intCol = intSourceCol To intSourceColSel
            'ԭ����ֵ,���߸��Ƶ�Ԫ����ֵ,����д�޸ı�־
            If intCol <> mlngOper Then
                If vsf.TextMatrix(intRow, intCol) <> "" Or arrRow(intCol - vsf.Col) <> "" Then
                    vsf.TextMatrix(intRow, intCol) = arrRow(intCol - vsf.Col)
                    vsf.Cell(flexcpData, intRow, intCol) = 1
                    vsf.RowData(intRow) = 1
                    mblnChange = True
                End If
            End If
        Next
    Next
    If mblnChange Then RaiseEvent AfterDataChanged
End Sub

Private Sub InitPanelMain()
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    
    dkpMain.SetCommandBars cbsThis
    
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockTopOf, Nothing): objPane.Title = "�༭": objPane.Options = PaneNoCaption
    objPane.Handle = picMain.hWnd
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    '63174:������,2013-07-03,������mblnEditable�ж��Ƿ���ز˵�ȡ�����ڲ˵���Update�¼��н��п��Ʋ˵��Ƿ�ɼ�.
    '��Ϊ�������ʱû��ѡ����mblnEditable=False,���ֲ˵�û�м��أ���ѡ����ʱmblnEditable=ture���˵������ڼ��ء�
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '�����
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "���"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�����Ŀ(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"):  cbrControl.ToolTipText = "ɾ����Ŀ(Alt+D)"
    End With
    
    '��ʾ����������
    '------------------------------------------------------------------------------------------------------------------
    Set objExtendedBar = cbsThis.Add("�鿴", xtpBarTop)
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    With objExtendedBar.Controls
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.Visible = mblnEditable
        pic����ȼ�.Visible = mblnEditable
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = pic����ȼ�.hWnd
        cbrCustom.ToolTipText = "����ȼ�"
        
        Set cbrControl = .Add(xtpControlLabel, 0, "��ʾ����")
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = txt��ʾ����.hWnd
        cbrCustom.ToolTipText = "��ʾ�������ڵ�����"
        Set cbrControl = .Add(xtpControlLabel, 0, "����")
        If Not mblnEditable Then cbrControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.Visible = mblnEditable
        picPati.Visible = mblnEditable
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = picPati.hWnd
        cbrCustom.ToolTipText = "�����б�"
    End With
    
    'Call SetDockRight(objExtendedBar, cbrToolBar)
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next

     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
        .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
        .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
    End With
    
    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function WriteIntoVsf(Optional ByRef strInfo As String) As Boolean
    Dim blnAllow As Boolean
    Dim StrText As String
    Dim strMsg As String
    Dim lngRecord As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long, strֵ�� As String
    
    If picInput.Visible Then
        lngRow = Split(txt����.Tag, "|")(0)
        lngCol = Split(txt����.Tag, "|")(1)
        If txt����.Enabled Then
            '������ݺϷ���
            If Val(cbo��λ.Tag) = 0 Then
                If txt����.Text <> "" Then
                    StrText = IIf(cbo��λ.Visible And Trim(cbo��λ.Text) <> "", cbo��λ.Text & ":", "") & Trim(txt����.Text)
                End If
            Else
                StrText = IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), cbo��λ.Text)
            End If
            If lngCol <= 2 Then
                If Trim(StrText) <> "" Then
                    strMsg = "Msgbox"
                    blnAllow = CheckDate2(lngRow, lngCol, StrText, strMsg)
                    strInfo = strMsg
                End If
            Else
                '��λ�ж�Ӧ�Ļ����¼���м��
                mrsSelItems.Filter = "��=" & lngCol
                mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
                
                intType = mrsItems!��Ŀ����     '0-��ֵ��1-����
                lngClass = mrsItems!��Ŀ����
                lngOrder = mrsItems!��Ŀ���
                strName = mrsItems!��Ŀ����
                lngLength = mrsItems!��Ŀ���� + IIf(NVL(mrsItems!��ĿС��, 0) = 0, 0, NVL(mrsItems!��ĿС��, 0) + 1)
                If intType = 0 Then
                    strֵ�� = NVL(mrsItems!��Ŀֵ��)
                Else
                    strֵ�� = ""
                    StrText = txt����.Text      '����������Ŀ,���û�ԭʼ¼��Ϊ׼
                End If
                
                '����Ǵ���ı��򲻼�����ݺϷ���
                If intType = 1 And lngLength > 100 Then
                    '�����κδ���
                    blnAllow = True
                Else
                    strMsg = "Msgbox"       '����ǿ�ֵ,��ʾ�������,���ڸñ����з��ش�����Ϣ
                    blnAllow = CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, strֵ��, strMsg)
                    strInfo = strMsg
                End If
                
                mrsItems.Filter = 0
                mrsSelItems.Filter = 0
            End If
            
            If blnAllow Then vsf.TextMatrix(lngRow, lngCol) = StrText
        Else
            blnAllow = True
            vsf.TextMatrix(lngRow, lngCol) = txt����.Text
        End If
    Else
        blnAllow = True
        lngRow = Split(lvwMultiSel.Tag, "|")(0)
        lngCol = Split(lvwMultiSel.Tag, "|")(1)
        vsf.TextMatrix(lngRow, lngCol) = strInfo
    End If
    txt����.Tag = ""
    cbo��λ.Visible = False
    txt����.Height = picInput.Height
    picInput.Visible = False
    lvwMultiSel.Visible = False
    
    '�����޸ı�־
    If blnAllow Then
        If picInput.Tag <> vsf.TextMatrix(lngRow, lngCol) Then
            '������޸ĵ�ʱ��,��Ҫ�Ѽ�¼ID��ͬ�����м�¼��ʱ��ȫ���޸���
            If lngCol <= 2 And Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 Then
                lngRows = vsf.Rows - 1
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                For lngRow = 1 To lngRows
                    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = lngRecord Then
                        vsf.TextMatrix(lngRow, lngCol) = StrText
                        '�޸ı�־
                        vsf.RowData(lngRow) = 1
                        vsf.Cell(flexcpData, lngRow, lngCol) = 1
                    End If
                Next
            Else
                '�޸ı�־
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                vsf.RowData(lngRow) = 1
                vsf.Cell(flexcpData, lngRow, lngCol) = 1
            End If
            mblnChange = True
        End If
        
        WriteIntoVsf = True
        If mblnChange Then RaiseEvent AfterDataChanged
    End If
End Function

Private Sub lvwMultiSel_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strData As String
    Dim intCol As Integer, intMax As Integer
    Dim blnAllow As Boolean
    
    If KeyCode = vbKeyReturn Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
        Call vsf_KeyDown(vbKeyReturn, Shift)
'    ElseIf KeyCode = vbKeyLeft Then
'        Call vsf_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    vsf.Width = picMain.Width
    vsf.Height = picMain.Height - vsf.Top
End Sub

Private Sub cbo��λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call txt����_KeyDown(vbKeyReturn, 0): Exit Sub
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrText As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode = vbKeyDown And InStr(1, "����������������", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0 Then
        If Shift = 0 Then
            cbo��λ.Tag = 0
            cbo��λ.Clear
            If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
                cbo��λ.AddItem "Ҹ��"
                cbo��λ.AddItem "����"
                cbo��λ.AddItem "����"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
                cbo��λ.AddItem ""
                cbo��λ.AddItem "����"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
                cbo��λ.AddItem "��������"
                cbo��λ.AddItem "������"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
                cbo��λ.AddItem "����"
                cbo��λ.AddItem "����"
                cbo��λ.AddItem "��������"
            End If
            If cbo��λ.ListCount <> 0 Then cbo��λ.ListIndex = 0
            cmdδ��˵��.ToolTipText = IIf(Val(cbo��λ.Tag) = 0, "�л���δ��˵��", "�л�����λ")
        ElseIf Shift = vbShiftMask Then
            gstrSQL = " Select ���� From ��������˵�� Order by ����"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡδ��˵��")
            With rsTemp
                cbo��λ.Clear
                Do While Not .EOF
                    cbo��λ.AddItem !����
                    .MoveNext
                Loop
                cbo��λ.ListIndex = 0
                cbo��λ.Tag = 1
            End With
        End If
        
        With cbo��λ
            .Top = picInput.Height - .Height
            .Width = picInput.Width
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        txt����.Height = picInput.Height - cbo��λ.Height
        If cbo��λ.Tag = 1 Then txt����.Text = cbo��λ.Text
        cmdδ��˵��.ToolTipText = IIf(Val(cbo��λ.Tag) = 0, "�л���δ��˵��", "�л�����λ")
    ElseIf KeyCode = vbKeyReturn Then
        Dim strData As String
        Dim lngCol As Long
        Dim blnAllow As Boolean
        
        blnAllow = True
        If Shift = vbCtrlMask Then Exit Sub
        If picInput.Visible And txt����.Tag <> "" Then
            lngCol = Split(txt����.Tag, "|")(1)
            If InStr(1, "������������", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
                '������ݺϷ���
                If cbo��λ.Tag = 0 Then
                    If txt����.Text <> "" Then
                        strData = IIf(cbo��λ.Visible And Trim(cbo��λ.Text) <> "", cbo��λ.Text & ":", "") & Trim(txt����.Text)
                    End If
                Else
                    strData = IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), cbo��λ.Text)
                End If
            Else
'                mrsSelItems.Filter = "��=" & lngCol
'                If mrsSelItems.RecordCount <> 0 Then
'                    mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
'                    If mrsItems.RecordCount <> 0 Then
'                        If mrsItems!��Ŀ���� = 1 Then
'                            strData = txt����.Text
'                        Else
                            strData = Trim(txt����.Text)
'                        End If
'                    End If
'                End If
'                mrsSelItems.Filter = 0
'                mrsItems.Filter = 0
            End If
            If strData <> picInput.Tag Then blnAllow = WriteIntoVsf(strData)
        End If
        
        If blnAllow Then
            Call vsf_KeyDown(vbKeyReturn, Shift)
        Else
            Call Vsf_EnterCell
            RaiseEvent AfterRowColChange(strData)
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If txt����.SelStart = 0 Then Call vsf_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        If Not (cmdδ��˵��.Visible And cbo��λ.Visible = False) Then Exit Sub
        StrText = frmWordsEditor.ShowMe(Me, mlng����ID, mlng��ҳID, txt����.Text)
        If StrText = "" Then Exit Sub
        txt����.Text = StrText

        DoEvents
        txt����.SetFocus
        Call txt����_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub txt��ʾ����_GotFocus()
    Call zlControl.TxtSelAll(txt��ʾ����)
End Sub

Private Sub txt��ʾ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then Call txt��ʾ����_Validate(blnCancel)
End Sub

Private Sub txt��ʾ����_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt��ʾ����_Validate(Cancel As Boolean)
    If Val(txt��ʾ����.Text) = Val(txt��ʾ����.Tag) Then Exit Sub
    txt��ʾ����.Tag = txt��ʾ����.Text
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), cbo����ȼ�.ListIndex, mstrPrivs, False, mblnEditable)
End Sub

Private Sub UserControl_GotFocus()
    Call Vsf_EnterCell
End Sub

Private Sub UserControl_Initialize()
    mstrSel = ""
    mstrSelItems = ""
    mblnShow = False
    mblnChange = False
    mblnInit = False
    txt��ʾ����.Tag = 1
    
    With cbo����ȼ�
        .Clear
        .AddItem "�ؼ�����ģ��"
        .AddItem "һ������ģ��"
        .AddItem "��������ģ��"
        .AddItem "��������ģ��"
    End With
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If InStr(1, "TXT����,CBO��λ", UCase(ActiveControl.Name)) <> 0 Then
            mblnShow = False
            picInput.Visible = False
            vsf.SetFocus
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    If mlng����ID = 0 Then
        picNothing.Visible = True
        picNothing.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        lblNothing.Move UserControl.ScaleWidth / 2 - lblNothing.Width / 2, UserControl.ScaleHeight / 2 - lblNothing.Height
    Else
        picNothing.Visible = False
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                strInfo = ""
            End If
            
            If mrsSelItems!��Ŀ��� = 1 Then
                strInfo = strInfo & Space(5) & "�����±�ʾ��:39/37.5"
            ElseIf mrsSelItems!��Ŀ��� = 3 Then
                If mbln���� = False Then strInfo = strInfo & Space(5) & "������׾��ʾ��:130/120"
            ElseIf vsf.TextMatrix(0, NewCol) Like "Ѫѹ*" Then
                strInfo = strInfo & Space(5) & "¼�����:����ѹ/����ѹ"
            End If
            
            If mrsSelItems!��Ŀ��� >= 1 And mrsSelItems!��Ŀ��� <= 3 Then
                strInfo = strInfo & Space(5) & "�������в�λѡ��;��SHIFT+������δ��˵����ѡ��"
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    RaiseEvent AfterRowColChange(strInfo)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    blnScroll = True
    Call Vsf_EnterCell
    blnScroll = False
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Vsf_EnterCell
End Sub

Private Sub vsf_DblClick()
    Dim blnDo As Boolean
    Dim lngOrder As Long
    
    If mblnEditable Then
        mblnShow = True
        Call Vsf_EnterCell
    Else
        If vsf.Row = 0 Then Exit Sub
        
        mrsSelItems.Filter = "��=" & vsf.Col
        blnDo = (mrsSelItems.RecordCount <> 0)
        If blnDo Then lngOrder = mrsSelItems!��Ŀ���
        mrsSelItems.Filter = 0
        If blnDo Then RaiseEvent DbClick(lngOrder & "|" & vsf.TextMatrix(vsf.Row, vsf.Col))
    End If
End Sub

Private Sub Vsf_EnterCell()
    Dim arrData
    Dim strData As String
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnAllow As Boolean, blnWords As Boolean
    Dim intCol As Integer, intMax As Integer
    
    If mblnInit = False Then Exit Sub
    Call ShowSignMarker
    
    '�����¼�������򱣴�
    blnAllow = True
    If picInput.Visible And txt����.Tag <> "" Then
        lngRow = Split(txt����.Tag, "|")(0)
        lngCol = Split(txt����.Tag, "|")(1)
        If InStr(1, "������������", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
            '������ݺϷ���
            If cbo��λ.Tag = 0 Then
                If txt����.Text <> "" Then
                    strData = IIf(cbo��λ.Visible And Trim(cbo��λ.Text) <> "", cbo��λ.Text & ":", "") & txt����.Text
                End If
            Else
                strData = IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), cbo��λ.Text)
            End If
        Else
'            mrsSelItems.Filter = "��=" & lngCol
'            If mrsSelItems.RecordCount <> 0 Then
'                mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
'                If mrsItems.RecordCount <> 0 Then
'                    If mrsItems!��Ŀ���� = 1 Then
'                        strData = txt����.Text
'                    Else
                        strData = Trim(txt����.Text)
'                    End If
'                End If
'            End If
'            mrsSelItems.Filter = 0
'            mrsItems.Filter = 0
        End If
        If strData <> picInput.Tag Then blnAllow = WriteIntoVsf(strData)
    ElseIf lvwMultiSel.Visible Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
    End If
    Call vsf.AutoSize(0, vsf.Cols - 1)
    picInput.Visible = False
    lvwMultiSel.Visible = False
    If blnAllow = False Then
        If vsf.Row <> lngRow Then vsf.Row = lngRow
        If vsf.Col <> lngCol Then vsf.Col = lngCol
        RaiseEvent AfterRowColChange(strData)
        Exit Sub
    End If
    
    RaiseEvent AfterSelChange(IIf(Trim(vsf.TextMatrix(vsf.Row, mlngSigner)) <> "", 1, 0), vsf.TextMatrix(vsf.Row, mlngCertLevel))
    
    mblnCheckVersion = CheckVersion
    If InStr(1, mstrPrivs, "�����¼�Ǽ�") = 0 Then Exit Sub
    If mblnShow = False Or IsPigeonhole Or Not mblnEditable Then Exit Sub
    If vsf.Col = 0 Or vsf.Row = 0 Then Exit Sub
    If vsf.Col = mlngOper And mblnCheckVersion = False Then Exit Sub
    If vsf.Col >= mlngSigner Then Exit Sub          'ǩ����,ǩ��ʱ���Լ���Ų�����༭,�������
    If vsf.RowIsVisible(vsf.Row) = False Then Exit Sub
    If Not blnScroll And vsf.Visible Then vsf.SetFocus
    
    '׼����ʾ
    With picInput
        .Tag = vsf.TextMatrix(vsf.Row, vsf.Col)             '����༭ǰ������
        .Left = vsf.ColPos(vsf.Col) + vsf.Left
        .Top = vsf.RowPos(vsf.Row) + vsf.Top
        .Width = vsf.ColWidth(vsf.Col)
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        If vsf.Row = vsf.Rows - 1 Then
            .Height = vsf.ROWHEIGHT(vsf.Row)    'ȡ���и�
        Else
            .Height = vsf.RowPos(vsf.Row + 1) - vsf.RowPos(vsf.Row)
        End If
        If .Height > vsf.RowHeightMax Then .Height = vsf.RowHeightMax
        If .Height < vsf.RowHeightMin Then .Height = vsf.RowHeightMin
        .ZOrder 0
        .Visible = True
    End With
    With cbo��λ
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        .Visible = False
        .Clear
        .Tag = 0
        blnAllow = True
        If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
            .AddItem "Ҹ��"
            .AddItem "����"
            .AddItem "����"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
            .AddItem ""
            .AddItem "����"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
            .AddItem "��������"
            .AddItem "������"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "����" Then
            .AddItem "����"
            .AddItem "����"
            .AddItem "��������"
            .Visible = True
            blnAllow = False
        Else
            '��λ��,����ǵ�ѡ,��ֵ�����������
            mrsSelItems.Filter = "��=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!��Ŀ��ʾ = 2 Then
                        '��ѡ
                        .AddItem " "
                        arrData = Split(NVL(mrsItems!��Ŀֵ��), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            If Mid(arrData(intCol), 1, 1) = "��" Then intIndex = intCol
                            .AddItem Replace(arrData(intCol), "��", "")
                        Next
                        blnAllow = False
                    ElseIf mrsItems!��Ŀ��ʾ = 3 Then
                        '��ѡ
                        picInput.Visible = False
                        lvwMultiSel.Font.Name = vsf.FontName
                        lvwMultiSel.Font.Size = vsf.FontSize
                        lvwMultiSel.Left = picInput.Left + picInput.Width - lvwMultiSel.Width
                        lvwMultiSel.Top = picInput.Top + picInput.Height
                        lvwMultiSel.Visible = True
                        If lvwMultiSel.Top + lvwMultiSel.Height > picMain.Height Then lvwMultiSel.Top = picInput.Top - lvwMultiSel.Height
                        
                        '��������
                        lvwMultiSel.ListItems.Clear
                        arrData = Split(NVL(mrsItems!��Ŀֵ��), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            strData = Replace(arrData(intCol), "��", "")
                            lvwMultiSel.ListItems.Add , "K" & intCol, strData
                            If Mid(arrData(intCol), 1, 1) = "��" Then lvwMultiSel.ListItems(intCol + 1).Selected = True
                            If InStr(1, "," & vsf.TextMatrix(vsf.Row, vsf.Col) & ",", "," & strData & ",") <> 0 Then lvwMultiSel.ListItems(intCol + 1).Checked = True
                        Next
                        lvwMultiSel.Tag = vsf.Row & "|" & vsf.Col
                        lvwMultiSel.SetFocus
                    ElseIf mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� >= 200 Then
                        blnWords = True
                    End If
                End If
            End If
            mrsSelItems.Filter = 0
            mrsItems.Filter = 0
        End If
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    With txt����
        .Enabled = blnAllow          '�����ǰ����������,������¼��
        .Text = vsf.TextMatrix(vsf.Row, vsf.Col)
        If .Enabled Then
            If InStr(1, .Text, ":") <> 0 And cbo��λ.ListCount Then
                With cbo��λ
                    If InStr(1, txt����.Text, ":") <> 0 Then
                        .Text = Split(txt����.Text, ":")(0)
                    End If
                    '.Top = picInput.Height - .Height
                    .Width = picInput.Width
                    .Visible = True
                    .FontName = vsf.FontName
                    .FontSize = vsf.FontSize
                    .ZOrder 0
                End With
                .Text = Split(.Text, ":")(1)
            End If
        Else
            If .Text <> "" Then cbo��λ.Text = .Text
            'If .Text = "" Then .Text = cbo��λ.Text
            With cbo��λ
                '.Top = picInput.Height - .Height
                .Width = picInput.Width
                .Visible = True
                .FontName = vsf.FontName
                .FontSize = vsf.FontSize
                .ZOrder 0
            End With
        End If
        .FontName = vsf.FontName
        .FontSize = vsf.FontSize
        .Width = picInput.Width
        .Height = picInput.Height - IIf(cbo��λ.Visible, cbo��λ.Height, 0)
        .Tag = vsf.Row & "|" & vsf.Col
    End With
    If cbo��λ.Enabled Then
        cbo��λ.Top = picInput.Height - cbo��λ.Height
        cbo��λ.Width = txt����.Width
    End If
    
    cmdδ��˵��.Visible = (InStr(1, "������������", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0) Or blnWords
    If cmdδ��˵��.Visible Then
        cmdδ��˵��.FontName = vsf.FontName
        cmdδ��˵��.FontSize = vsf.FontSize
        '���������������Ŀ,���¼������ݲ�����ֵ��,�򽫱�־��Ϊ1
        If InStr(1, txt����.Text, "/") = 0 Then
            If Trim(Split(txt����.Text & "|", "|")(0)) <> "" And Trim(Split(txt����.Text & "|", "|")(0)) <> "����" Then
                If Not IsNumeric(Split(txt����.Text & "|", "|")(0)) Then
                    strData = Split(txt����.Text & "|", "|")(0)
                    Call txt����_KeyDown(vbKeyDown, vbShiftMask)
                    txt����.Text = strData
                End If
            End If
        End If
        If blnWords Then
            cmdδ��˵��.ToolTipText = "���԰�Ctrl+W�����ʾ�༭��"
        Else
            cmdδ��˵��.ToolTipText = IIf(Val(cbo��λ.Tag) = 0, "�л���δ��˵��", "�л�����λ")
        End If
        cmdδ��˵��.Left = txt����.Width - cmdδ��˵��.Width
    End If
    
    On Error Resume Next
    If txt����.Enabled Then
        txt����.SetFocus
    Else
        cbo��λ.SetFocus
    End If
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intStep As Integer
    
    '�������������,�Ե�
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 _
        Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then    'Or KeyCode = vbKeyLeft
        Exit Sub
    End If
    If KeyCode = vbKeyLeft And (picInput.Visible = False And lvwMultiSel.Visible = False) Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        '�����ǰ��Ԫ�������
        vsf.TextMatrix(vsf.Row, vsf.Col) = ""
        cbo��λ.Visible = False
        txt����.Text = ""
        txt����.Height = picInput.Height
    End If
    
    If KeyCode = vbKeyReturn Then
        '������һ����Ч��Ԫ��
toNextCol:
        If vsf.Col < mlngSigner Then
            vsf.Col = vsf.Col + 1
            If vsf.Col = mlngSigner Then GoTo toNextCol
            If vsf.ColHidden(vsf.Col) Then GoTo toNextCol
        Else
toNextRow:
            If vsf.Row = vsf.Rows - 1 Then
                vsf.Rows = vsf.Rows + 1
                vsf.TextMatrix(vsf.Rows - 1, 1) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                vsf.TextMatrix(vsf.Rows - 1, 2) = Format(zlDatabase.Currentdate, "HH:mm")
                'Vsf.Cell(flexcpAlignment, Vsf.Rows - 1, 0, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
            End If
            vsf.Row = vsf.Row + 1
            If vsf.RowHidden(vsf.Row) Then GoTo toNextRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.Row) = False Then
            vsf.TopRow = vsf.Row
        End If
        Exit Sub
    End If
    
    If KeyCode = vbKeyLeft Then
        '������һ����Ч��Ԫ��
toPreCol:
        If vsf.Col > 1 Then
            vsf.Col = vsf.Col - 1
            If vsf.Col >= mlngSigner Then GoTo toPreCol
            If vsf.Col = mlngOper Then GoTo toPreCol
            If vsf.ColHidden(vsf.Col) Then GoTo toPreCol
        Else
toPreRow:
            If vsf.Row > 1 Then
                vsf.Row = vsf.Row - 1
                vsf.Col = vsf.Cols - 1
                GoTo toPreCol
            Else
                vsf.Row = 1
            End If
            If vsf.RowHidden(vsf.Row) Then GoTo toPreRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.Row) = False Then
            vsf.TopRow = vsf.Row
        End If
        Exit Sub
    End If
    
    mblnShow = True
    Call Vsf_EnterCell
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer, intCol As Integer
    If Button <> 1 Then Exit Sub
    
    intCol = vsf.MouseCol
    intRow = vsf.MouseRow
    If intRow = 0 And intCol = 0 Then
        Call vsf.Select(0, 0, vsf.Rows - 1, vsf.Cols - 1)
    ElseIf intCol = 0 Then
        Call vsf.Select(intRow, 0, intRow, vsf.Cols - 1)
    ElseIf intRow = 0 Then
        Call vsf.Select(0, intCol, vsf.Rows - 1, intCol)
    End If
End Sub

Public Sub ArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or mblnMoved_HL Then Exit Sub
    If MsgBox("��Ҫ���ò��˱���סԺ���л����¼�鵵��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "Zl_���ӻ����¼_Archive(" & mlng����ID & "," & mlng��ҳID & "," & cbo����.ItemData(cbo����.ListIndex) & ",'" & gstrUserName & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�鵵")

        mstrPigeonhole = gstrUserName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or mblnMoved_HL Then Exit Sub
    If mstrPigeonhole <> "" Then
        If MsgBox("��Ҫ�����ò��˱���סԺ�����ѹ鵵�����¼��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

            gstrSQL = "Zl_���ӻ����¼_UnArchive(" & mlng����ID & "," & mlng��ҳID & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����鵵")
            mstrPigeonhole = ""
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SignMe()
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim strTime As String
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim intRow As Integer, intRows As Integer
    On Error GoTo errHand
    '������ʱ��ѭ������ǩ��
    
    If mlng����ID = 0 Or mblnMoved_HL Then Exit Sub
    
    intRows = vsf.Rows - 1
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For intRow = 1 To intRows
        If vsf.TextMatrix(intRow, mlngSigner) = "" And vsf.TextMatrix(intRow, vsf.Cols - 1) = gstrUserName Then
            If strTime <> vsf.TextMatrix(intRow, 1) & " " & vsf.TextMatrix(intRow, 2) & ":00" And Val(vsf.TextMatrix(intRow, mlngRecord)) <> 0 Then
                strTime = vsf.TextMatrix(intRow, 1) & " " & vsf.TextMatrix(intRow, 2) & ":00"
                If SignName(strTime, strSignTime, str״̬) = False Then Exit For
                blnSign = True
            End If
        End If
    Next
    If blnSign Then Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), cbo����ȼ�.ListIndex, mstrPrivs, False, mblnEditable)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnSignMe()
    Dim blnUnSign As Boolean
    Dim strTime As String               '��¼ʱ��
    Dim strSignTime As String           'ǩ��ʱ��
    Dim intRow As Integer, intRows As Integer
    Dim lng��ֹ�汾 As Long             '���汾
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������ȡ��ǩ��,ֻȡ����ǰѡ��ļ�¼
    
    If mlng����ID = 0 Or mblnMoved_HL Then Exit Sub
    
    If vsf.TextMatrix(vsf.Row, mlngSigner) <> gstrUserName Then
        MsgBox "������ȡ����������Ա��ǩ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    '��ͬһǩ��ʱ���������ȡ����,����ȡ��ǩ��
'    strSignTime = Vsf.TextMatrix(Vsf.Row, mlngSignTime)
'    gstrSQL = " Select A.����ʱ�� From ���˻����¼ A,���˻������� B" & _
'              " Where A.ID=B.��¼ID And B.��¼����=5 And B.��Ŀ����=[4]" & _
'              " And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And A.������Դ=2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǩ��", mlng����id, mlng��ҳid, cbo����.itemdata(cbo����.listindex), strSignTime)
'    With rsTemp
'        Do While Not .EOF
'            If UnSignName(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), blnClear) = False Then Exit Sub
'            blnUnSign = True
'            .MoveNext
'        Loop
'    End With
    
    If UnSignName(vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00", blnClear) = False Then Exit Sub
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), cbo����ȼ�.ListIndex, mstrPrivs, False, mblnEditable)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, str״̬ As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strSource As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5 And Nvl(a.��ʼ�汾,1)=Nvl(b.���汾,1)"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�Ƿ��Ѿ�ǩ����", mlng����ID, mlng��ҳID, CDate(strStart), cbo����.ItemData(cbo����.ListIndex))
    If rs.BOF = False Then
        MsgBox "��ǰû����Ҫǩ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
        
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.�޸�ʱ��" & vbNewLine & _
             " From ���˻������� a,���˻����¼ b " & vbNewLine & _
             " Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��ֹ�汾 Is Null" & vbNewLine & _
             " Order by A.��Ŀ���"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", mlng����ID, mlng��ҳID, CDate(strStart), cbo����.ItemData(cbo����.ListIndex))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    Debug.Print "ǩ����" & Now & vbCrLf & strSource
    If strSource = "" Then
        MsgBox "��ǰû����Ҫǩ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '76223:������,2014-08-05,����ǩ�����ʱ�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, mlng����ID, mlng��ҳID, mlng����ID, str״̬)
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���ӻ����¼_SignName("
        gstrSQL = gstrSQL & mlng����ID & "," & mlng��ҳID & "," & cbo����.ItemData(cbo����.ListIndex) & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "',"
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "','" & oSign.ʱ�����Ϣ & "')"

        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UnSignName(ByVal strStart As String, ByVal blnClear As Boolean) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim strSource As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.��ĿID AS ֤��ID From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�Ƿ��Ѿ�ǩ����", mlng����ID, mlng��ҳID, CDate(strStart), cbo����.ItemData(cbo����.ListIndex))
    If rs.BOF Then
        MsgBox "��ǰû����Ҫȡ����ǩ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����ǵ���ǩ��,����Ҫ��֤
    '------------------------------------------------------------------------------------------------------------------
    If Val(NVL(rs!֤��ID, 0)) > 0 Then
        '����ǩ����֤
        Err.Clear
        If gobjTendESign Is Nothing Then
            On Error Resume Next
            Set gobjTendESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear
            On Error GoTo 0
            If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjTendESign Is Nothing Then
            If Not gobjTendESign.CheckCertificate(gstrDBUser) Then Exit Function
        Else
            MsgBox "����ǩ������δ����ȷ��װ�����˲������ܼ�����", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Zl_���ӻ����¼_Unsignname("
    gstrSQL = gstrSQL & mlng����ID & ","
    gstrSQL = gstrSQL & mlng��ҳID & ","
    gstrSQL = gstrSQL & cbo����.ItemData(cbo����.ListIndex) & ","
    gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & _
                      IIf(blnClear, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ȡ��ǩ��")
    
    UnSignName = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), mbyt����ȼ�, mstrPrivs, True, mblnEditable)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    mblnShow = False
    picInput.Visible = False
    
    SaveME = True
    
    Call ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, cbo����.ItemData(cbo����.ListIndex), mbyt����ȼ�, mstrPrivs, False, mblnEditable)
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngPatiID As Long, ByVal lngPageId As Long, lngDeptId As Long, _
    Optional ByVal intBaby As Integer = 0, Optional ByVal byt������ As Byte = 3, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnEditable As Boolean = True)
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '���أ� ��
    '******************************************************************************************************************
'    Dim bln������ As Boolean
    
    Err = 0
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mblnInit = False
    mblnEditable = blnEditable And Not mblnMoved_HL
    
    lngLastRow = vsf.Row
    lngLastTopRow = vsf.TopRow
    lngLastPatientID = mlng����ID
    If lngLastRow < 1 Then lngLastRow = 1
    If lngLastTopRow < 1 Then lngLastTopRow = 1
    
    If mblnChange And Not blnCancel Then
        If MsgBox("��ǰ���˵����ݻ�δ���棬�㡰�ǡ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
    mblnShow = False
    picInput.Visible = False
    
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mlng����ID = lngDeptId
    mintӤ�� = intBaby
    mbyt����ȼ� = byt������
    mstrPrivs = strPrivs
    Set mfrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitPanelMain
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    '��ȡ�ò�����������,�Ա������ȡģ��
    Call UserControl_Resize
    If mlng����ID = 0 Then Exit Sub
    gstrSQL = " Select ��Ժ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������", mlng����ID, mlng��ҳID)
    mlng����ID = rsTemp!��Ժ����ID
    
    '�����˷����仯ʱ,������±���
    If lngLastPatientID <> mlng����ID Then
        mstrSel = ""
        mstrSelItems = ""
        cbo����ȼ�.ListIndex = mbyt����ȼ�
        
        '��ȡ���˵�Ӥ��
        gstrSQL = " Select NVL(A.Ӥ������,NVL(C.����,B.����) ||'֮��'||A.���) AS ����,A.���" & _
                  " From ������Ϣ B,������ҳ C,������������¼ A" & _
                  " Where B.����ID=C.����ID And A.����ID=C.����ID And A.��ҳID=C.��ҳID And C.����ID=[1] And C.��ҳID=[2]" & _
                  " Order By A.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�Ӥ��", mlng����ID, mlng��ҳID)
        With cbo����
            .Clear
            .AddItem "���˱���"
            
            Do While Not rsTemp.EOF
                .AddItem rsTemp!����
                .ItemData(.NewIndex) = rsTemp!���
                rsTemp.MoveNext
            Loop
        End With
    End If
    cbo����.ListIndex = mintӤ��
    
    Call InitBill
    Call ReadData
    mblnInit = True
    
    '�ָ���λ
    If lngLastPatientID <> mlng����ID Then
        lngLastRow = 1
        lngLastTopRow = 1
    End If
    
    cbo����.Tag = cbo����.ListIndex
    If vsf.Rows - 1 > lngLastRow Then vsf.Row = lngLastRow
    If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = lngLastTopRow
    Call Vsf_EnterCell
    Call ReSetFontSize(mbytFontSize)
    mblnChange = False
    RaiseEvent AfterRefresh
    
    'Call OutputRsData(mrsSelItems)
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckData() As Boolean
    Dim StrText As String
    Dim strMaxDate As String, strֵ�� As String
    Dim lngRow As Long, lngRows As Long, lngCol As Long
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long
    On Error GoTo errHand
    '�������¼��Ϸ���
    
    lngRows = vsf.Rows - 1
    '�ȼ�������Ƿ�Ϸ�
    For lngRow = 1 To lngRows
        If Val(vsf.RowData(lngRow)) = 1 Then
            If Not CheckDate1(lngRow) Then
                vsf.Row = lngRow
                vsf.Col = 1
                If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = vsf.Row
                Exit Function
            End If
        End If
    Next
    
    '���μ�������Ŀ��¼��Ϸ���
    With mrsSelItems
        .MoveFirst
        Do While Not .EOF
            mrsItems.Filter = "��Ŀ���=" & !��Ŀ���
            If mrsItems.RecordCount <> 0 Then
                lngCol = !��
                intType = mrsItems!��Ŀ����     '0-��ֵ��1-����
                lngClass = mrsItems!��Ŀ����
                lngOrder = mrsItems!��Ŀ���
                strName = mrsItems!��Ŀ����
                lngLength = mrsItems!��Ŀ���� + IIf(NVL(mrsItems!��ĿС��, 0) = 0, 0, NVL(mrsItems!��ĿС��, 0) + 1)
                If intType = 0 Then
                    strֵ�� = NVL(mrsItems!��Ŀֵ��)
                Else
                    strֵ�� = ""
                End If
                '��ֵ��Ŀ:ֻ������,����������,�Լ�Ѫѹ�Ŵ���/¼��
                '�ı���Ŀ:ֻ����Ƿ񳬳�
                If Not (intType = 1 And lngLength > 100) Then
                    For lngRow = 1 To lngRows
                        If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                            StrText = vsf.TextMatrix(lngRow, lngCol)
                            If Trim(StrText) <> "" Then
                                If Not CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, strֵ��) Then
                                    vsf.Row = lngRow
                                    If vsf.RowIsVisible(vsf.Row) Then vsf.TopRow = vsf.Row
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    mrsItems.Filter = 0
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
End Function

Private Function CheckDate1(ByVal lngRow As Long) As Boolean
    If Not IsDate(vsf.TextMatrix(lngRow, 1)) Then
        MsgBox "���ڸ�ʽ����yyyy-MM-dd", vbInformation, gstrSysName
        Exit Function
    End If
    If vsf.TextMatrix(lngRow, 1) > mstrMaxDate Then
        MsgBox "��" & lngRow & "�е����ڴ����˲���[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Trim(vsf.TextMatrix(lngRow, 2)) = "" Then
        MsgBox "��" & lngRow & "�е�ʱ�䲻��Ϊ�գ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Len(vsf.TextMatrix(lngRow, 2)) = 2 Then vsf.TextMatrix(lngRow, 2) = vsf.TextMatrix(lngRow, 2) & ":00"
    
    CheckDate1 = True
End Function

Private Function CheckDate2(ByVal lngRow As Long, ByVal lngCol As Long, StrText As String, Optional ByRef strInfo As String = "") As Boolean
    Dim strMsg As String
    Dim strDate As String
    Dim rsTemp As New ADODB.Recordset
    
    '����С����Ժ����,ʱ�䲻�㲹λʱ,Ҫ�����Ƿ�Ϸ�
    If lngCol = 1 Then
        gstrSQL = " Select ��Ժ���� From ������ҳ Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ��Ժ����")
        strDate = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
        
        If Not IsDate(StrText) Then
            strMsg = "���ڸ�ʽ����yyyy-MM-dd"
            GoTo errHand
        End If
        If StrText < strDate Then
            strMsg = "��" & lngRow & "�е�����С������Ժ���ڣ�"
            GoTo errHand
        End If
        If StrText > mstrMaxDate Then
            strMsg = "��" & lngRow & "�е����ڴ����˲���[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            GoTo errHand
        End If
    End If
    
    If lngCol = 2 Then
        If Trim(StrText) = "" Then
            strMsg = "��" & lngRow & "�е�ʱ�䲻��Ϊ�գ�"
            GoTo errHand
        End If
        If Len(StrText) <= 2 Then StrText = String(2 - Len(StrText), "0") & StrText
        If Val(Mid(StrText, 1, 2)) < 0 Or Val(Mid(StrText, 1, 2)) > 23 Then
            strMsg = "��" & lngRow & "�е�ʱ�����ݷǷ���"
            GoTo errHand
        End If
        If Len(StrText) = 2 Then StrText = StrText & ":00"
        If Len(StrText) < 5 And InStr(1, StrText, ":") > 0 Then StrText = String(5 - Len(StrText), "0") & StrText
        If Mid(StrText, 3, 1) <> ":" Then
            strMsg = "��" & lngRow & "�е�ʱ�����ݸ�ʽ�Ƿ�[09:00]��"
            GoTo errHand
        End If
        If Len(StrText) < 5 Then StrText = StrText & String(5 - Len(StrText), "0")
        If Not (Val(Mid(StrText, 4, 2)) >= 0 And Val(Mid(StrText, 4, 2)) <= 59) Then
            strMsg = "��" & lngRow & "�е�ʱ�����ݸ�ʽ�Ƿ�[09:00]��"
            GoTo errHand
        End If
        vsf.TextMatrix(lngRow, 2) = StrText
    
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(lngRow, mlng����ID, mlng��ҳID, vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), strMsg) Then
            GoTo errHand
        End If
    End If
    CheckDate2 = True
    Exit Function
errHand:
    If strInfo = "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        strInfo = strMsg
    End If
End Function

Private Function CheckValid(ByRef StrText As String, ByVal lngOrder As Long, ByVal lngType As Long, ByVal strCap As String, _
    ByVal lngLength As Long, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal strֵ�� As String, _
    Optional ByRef strInfo As String = "") As Boolean
    Dim arrData
    Dim strMsg As String
    Dim intDo As Integer, intCount As Integer
    Dim strPart As String, strValue1 As String, strValue2 As String, strTextClone As String
    
    If StrText = "" Then
        CheckValid = True
        Exit Function
    End If
    
    '��ȡ����λ,��/������
    strTextClone = StrText
    If InStr(1, strTextClone, ":") <> 0 Then
        strPart = Split(strTextClone, ":")(0)
        strTextClone = Split(strTextClone, ":")(1)
    End If
    If InStr(1, strTextClone, "/") <> 0 Then
        strValue1 = Split(strTextClone, "/")(0)
        strValue2 = Split(strTextClone, "/")(1)
    Else
        strValue1 = strTextClone
    End If
    
    If lngType = 2 Then '����ǻ��Ŀ����ܴ��ڲ�λ,�Ѳ�λ�����,ֻ���¼��������Ƿ񳬹�����
        If InStr(1, StrText, ":") <> 0 Then
            StrText = Split(StrText, ":")(1)
        End If
    End If
    
'    If strֵ�� = "" Then  '��ͨ��Ŀ
'        If Not (lngOrder = 9 Or lngOrder = 10) Then '���������ų�����������Ч��Χ���
'            If LenB(StrConv(strText, vbFromUnicode)) > lngLength Then
'                strMsg = "��" & lngRow & "�е�" & strCap & "���������飡"
'                GoTo errHand
'            End If
'        End If
'    Else                    '�������������Լ�Ѫѹ
        'û�����ʵ�ʱ�򣬲�����¼������
        If lngOrder = 2 And mbln���� Then
            If InStr(1, StrText, "/") <> 0 Then
                strMsg = "�뽫��õ���������¼�뵥�������ʵ�Ԫ���У�"
                GoTo errHand
            End If
        End If
        If lngOrder = 3 Then
            If InStr(1, StrText, "/") <> 0 Then
                strMsg = "��������¼�����"
                GoTo errHand
            End If
        End If
        If lngOrder = 4 Or lngOrder = 5 Then
            'Ѫѹֵ���뺬/
            If vsf.TextMatrix(0, lngCol) Like "Ѫѹ*" Then
                If InStr(1, StrText, "/") = 0 Then
                    strMsg = "Ѫѹ���ݵĸ�ʽ��������ѹ/����ѹ��"
                    GoTo errHand
                End If
                If Trim(Split(StrText, "/")(0)) = "" Or Trim(Split(StrText, "/")(1)) = "" Then
                    strMsg = "Ѫѹ���ݴ�������ѹ/����ѹ��"
                    GoTo errHand
                End If
            End If
        End If
        If UBound(Split(StrText, "/")) > 1 Then
            strMsg = "��" & lngRow & "�е�" & strCap & "����¼��������飡"
            GoTo errHand
        End If
        
        arrData = Split(StrText, "/")
        intCount = UBound(arrData)
        For intDo = 0 To intCount
            StrText = arrData(intDo)
            If InStr(1, StrText, ":") <> 0 Then StrText = Split(StrText, ":")(1)
            '������Ŀ������Ƿ񳬳�
            If lngOrder > 3 Then
                If LenB(StrConv(StrText, vbFromUnicode)) > lngLength Then
                    strMsg = "��" & lngRow & "�е�" & strCap & "���������飡"
                    vsf.TopRow = lngRow
                    GoTo errHand
                End If
            End If
            If IsNumeric(StrText) Then    '��Ч��Χ�뵱ǰ¼��ֵ������ֵ�Ͳż��,���򵱳���δ��˵��
                If Not (lngOrder = 9 Or lngOrder = 10) Then '���������ų�����������Ч��Χ���
                    If strֵ�� <> "" Then
                        If IsNumeric(Split(strֵ��, ";")(0)) Then
                            If Not (Val(StrText) >= Split(strֵ��, ";")(0) And Val(StrText) <= Split(strֵ��, ";")(1)) Then
                                strMsg = "��" & lngRow & "�е�" & strCap & "������Ч��Χ��" & Split(strֵ��, ";")(0) & "-" & Split(strֵ��, ";")(1) & "�������飡"
                                GoTo errHand
                            End If
                        End If
                    End If
                    If mrsItems!��Ŀ���� = 0 Then
                        If NVL(mrsItems!��ĿС��, 0) <> 0 Then
                            If intDo = 0 Then
                                strValue1 = Format(StrText, "#0." & String(mrsItems!��ĿС��, "0"))
                            Else
                                strValue2 = Format(StrText, "#0." & String(mrsItems!��ĿС��, "0"))
                            End If
                        Else
                            If intDo = 0 Then
                                strValue1 = Format(StrText, "#0")
                            Else
                                strValue2 = Format(StrText, "#0")
                            End If
                        End If
                    End If
                End If
            End If
        Next
'    End If
    
    'ƴװ���봮
    StrText = IIf(strPart <> "", strPart & ":", "") & strValue1 & IIf(strValue2 <> "", "/" & strValue2, "")
    
    CheckValid = True
    Exit Function
errHand:
    If strInfo = "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        strInfo = strMsg    '���������Ϣ,�ɵ��ó�����
    End If
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, blnOper As Boolean         'ָ��ĳ��ʱ������Ƿ��������
    Dim lngOrder As Long
    Dim strTime As String, strTmp As String, strSQLtmp As String, strMsg As String
    Dim intAllow As Integer, intType As Integer, lngClass As Long
    Dim str���� As String, str��� As String, str��λ As String, strδ��˵�� As String 'str���:ֻ�������⽵�»�������׾
    Dim lngRecord As Long, lngGroup As Long, lngMAX As Long
    Dim lngRow As Long, lngRows As Long, lngCol As Long, lngCols As Long
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim intPos As Integer, intMax As Integer
    Dim strSQL() As String
    On Error GoTo errHand
    'ͬһ��ʱ����(ͬһ����¼ID),��������ֶ�������,Ҳ����ֻ����һ��������������Ĵ���
    '�����¼ID=0,�����ļ�¼,��ʱ����Ѵ�����ʷ��¼��,�����������ű���
    
    If mblnMoved_HL Then Exit Function
    
    ReDim Preserve strSQL(1 To 1)
    lngRows = vsf.Rows - 1
    lngCols = mlngSigner - 1         '�����ǩ����,ǩ��ʱ��,��¼ID,��Ų�����
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '1�����ʱ���б仯���ȴ���ʱ��
    For lngRow = 1 To lngRows
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        strMsg = "msgbox"
        If Val(vsf.RowData(lngRow)) = 1 Then
            If Not CheckTime(lngRow, mlng����ID, mlng��ҳID, vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2), Mid(strDate, 1, 16), strMsg) Then
                Exit Function
            End If
        End If
        
        If Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 And (vsf.Cell(flexcpData, lngRow, 1) = 1 Or vsf.Cell(flexcpData, lngRow, 2) = 1) Then
            If lngRecord <> Val(vsf.TextMatrix(lngRow, mlngRecord)) Then
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                strStart = vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2) & ":00"
                gstrSQL = "Zl_���˻����¼_UpdateReplace(" & lngRecord & ",0," & cbo����.ItemData(cbo����.ListIndex) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
        End If
    Next
    
    '2�������δ���༭����Ԫ��
    For lngRow = 1 To lngRows
        '�ȶ�λ�޸Ĺ�����,��������ѭ���ҵ��޸Ĺ�����
        If vsf.TextMatrix(lngRow, 1) <> "" And vsf.TextMatrix(lngRow, 2) <> "" Then
            If strTime <> vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2) Then
                strTime = vsf.TextMatrix(lngRow, 1) & " " & vsf.TextMatrix(lngRow, 2)
                blnOper = False
            End If
            
            strDate = strTime
            strStart = strDate & ":00"
            strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"
            
            If Val(vsf.RowData(lngRow)) = 1 Then
                '�������ȡ��ţ�����ţ���ȡ��ǰ������
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                lngGroup = Val(vsf.TextMatrix(lngRow, mlngGroup))
                '�п���ԭ���������е���Ų��ǰ�˳�����ӵ�,��˴˶ν���У��
                If lngGroup = 0 Then
                    'ȡ�������
                    gstrSQL = " select max(��¼���) AS ��� " & _
                              " From ���˻�������" & _
                              " where ��¼ID=(" & _
                              "     select ID from ���˻����¼" & _
                              "     where ����ID=[1] and ��ҳID=[2] and Ӥ��=[3] and ����ID=[4] and ����ʱ��=[5])"
                    If mblnMoved_HL Then
                        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
                        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", mlng����ID, mlng��ҳID, cbo����.ItemData(cbo����.ListIndex), mlng����ID, CDate(strStart))
                    lngGroup = NVL(rsTemp!���, 0) + 1
                End If
                
                'һ��Ԫ��һ��Ԫ�صĴ���
                For lngCol = 3 To lngCols
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        '�����ݽ������������޸Ĳ���
                        gstrSQL = "Zl_���˻����¼_UpdateRecord("
                        gstrSQL = gstrSQL & mlng����ID & "," & mlng��ҳID & "," & cbo����.ItemData(cbo����.ListIndex) & ","
                        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & IIf(lngCol <> mlngOper, 1, 4) & ","
                        
                        lngOrder = 0
                        If lngCol <> mlngOper Then
                            mrsSelItems.Filter = "��=" & lngCol
                            mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
                            intType = mrsItems!��Ŀ����
                            lngClass = mrsItems!��Ŀ����
                            lngOrder = mrsItems!��Ŀ���
                        End If
                        strSQLtmp = gstrSQL     'ΪѪѹ��������
                        gstrSQL = gstrSQL & lngOrder & ","
                        
                        str��λ = "": str��� = "": strδ��˵�� = ""
                        str���� = vsf.TextMatrix(lngRow, lngCol)
                        If (lngOrder = 1 Or lngOrder = 2 Or lngOrder = 3) Or lngClass = 2 Then
                            If InStr(1, str����, ":") <> 0 Then
                                str��λ = Trim(Split(str����, ":")(0))
                                str���� = Trim(Split(str����, ":")(1))
                            End If
                            If InStr(1, str����, "/") <> 0 Then
                                str��� = Trim(Split(str����, "/")(1))
                                str���� = Trim(Split(str����, "/")(0))
                            End If
                        ElseIf lngOrder = 4 Then        '��Ϊ�ǰ���ѭ��,����ֻ�ᴦ��һ��,����Ǻϲ�¼������ѹ������ѹ,���ڱ�����ٴ�����
                            If InStr(1, str����, "/") <> 0 Then
                                str���� = Split(str����, "/")(lngOrder - 4)
                            End If
                        End If
                        'ֻ��������Ŀ�Ŵ���δ��˵���ĸ���
                        If lngOrder <= 3 And Not IsNumeric(str����) And lngCol <> mlngOper Then
                            If (lngOrder = 1 And str���� <> "����") Or lngOrder <> 1 Then
                                strδ��˵�� = str����
                                str���� = ""
                            End If
                        End If
                        
                        '����������Ŀ,�����/��1
                        If lngOrder = -1 Then
                            gstrSQL = gstrSQL & "1,"
                        Else
                            gstrSQL = gstrSQL & "0,"
                        End If
                        
                        If lngCol <> mlngOper Or blnOper = False Then
                            gstrSQL = gstrSQL & "'" & str���� & "','" & str��λ & "'," & intAllow & "," & IIf(IsNumeric(str����), 0, 1) & "," & lngGroup & ",'" & strδ��˵�� & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                            
                            '�����Ѫѹ
                            If lngOrder = 4 And vsf.TextMatrix(0, lngCol) Like "Ѫѹ*" Then
                                If str���� <> "" Then str���� = Split(vsf.TextMatrix(lngRow, lngCol), "/")(1)       '��Ϊ��ʱ���и�ֵ,Ϊ����˵���������������
                                strSQLtmp = strSQLtmp & "5,0,"
                                gstrSQL = strSQLtmp & "'" & str���� & "','" & str��λ & "'," & intAllow & "," & IIf(IsNumeric(str����), 0, 1) & "," & lngGroup & ",'" & strδ��˵�� & "')"
                                strSQL(ReDimArray(strSQL)) = gstrSQL
                            End If
                            
                            If lngCol = mlngOper Then blnOper = True
                        End If
                        
                        '----------------------------------------------------------------------------
                        'û��ѡ������,����������������ͬʱ¼��(�����Ϊ��,��ɱ�ǲ�����������Ĺ���)
                        If (lngOrder = 1 Or lngOrder = 2 And mbln���� = False) Then
            
                            gstrSQL = "Zl_���˻����¼_UpdateRecord("
                            gstrSQL = gstrSQL & mlng����ID & "," & mlng��ҳID & "," & cbo����.ItemData(cbo����.ListIndex) & ","
                            gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "1,"
                            gstrSQL = gstrSQL & IIf(lngOrder = 2, -1, lngOrder) & ","
                            gstrSQL = gstrSQL & "1,"
                                                            
                            If str��� <> "" And str���� <> "" Then
                                Select Case intType
                                Case 0          '��ֵ
                                    strTmp = Val(str���)
                                Case 1          '�ı�
                                    strTmp = str���
                                End Select
                                gstrSQL = gstrSQL & "'" & strTmp & "','" & str��λ & "'," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & lngGroup & ",Null)"
                            Else
                                gstrSQL = gstrSQL & "NULL,'" & str��λ & "'," & intAllow & ",0," & lngGroup & ",Null)"
                            End If
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'ѭ��ִ��SQL��������
    gcnOracle.BeginTrans
    blnTrans = True
    intMax = UBound(strSQL)
    For intPos = 1 To intMax
        If strSQL(intPos) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intPos), "������������")
    Next
    SaveData = True
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnChange = False
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    
    RaiseEvent AfterDataChanged
    RaiseEvent AfterRefresh
    Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
End Function


'---------------------------------------------------------------------------------
'�����ǻ������������
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���,ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    With rsObj
        Do While Not .EOF
            Debug.Print !�� & "," & !��Ŀ��� & "," & !��Ŀ����
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function CheckVersion(Optional ByVal lngRow As Long = 0, Optional ByVal lngCol As Long = 0) As Boolean
    Dim lng��Ŀ��� As Long
    Dim lng��ǰ�汾 As Long
    Dim lng��߰汾 As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������ֻ����һ����¼,ֻ�е�ǩ����¼�����汾С���������ݵĿ�ʼ�汾ʱ,��������б༭(���������)
    '���Ҫ���һ��,���д���������¼,���������������н��б༭,��ȡ���ò���
    
    If lngRow = 0 Then lngRow = vsf.Row
    If lngCol = 0 Then lngCol = vsf.Col
    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = 0 Then CheckVersion = True: Exit Function      '�¼�¼ֱ���˳�
    If vsf.Cell(flexcpData, lngRow, lngCol) <> 0 Then CheckVersion = True: Exit Function                              '���������������������
    
    'ȡ��ǰ��Ԫ�����Ŀ���
    mrsSelItems.Filter = "��=" & lngCol
    If mrsSelItems.RecordCount <> 0 Then
        lng��Ŀ��� = mrsSelItems!��Ŀ���
    Else
        mrsSelItems.Filter = 0
        Exit Function
    End If
    mrsSelItems.Filter = 0
    
    'ȡ��ǰ��¼+��ŵ����汾
    gstrSQL = " Select Max(��ʼ�汾) AS ��߰汾 From ���˻������� Where ��¼ID=[1] And ��¼����=5"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ��¼+��ŵ����汾", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)))
    lng��߰汾 = NVL(rsTemp!��߰汾, 0)
    
    'ȡ��ǰ��Ŀ�ĵ�ǰ�汾
    gstrSQL = " Select MAX(��ʼ�汾) AS ��ǰ�汾 From ���˻������� Where ��¼ID=[1] And ��¼���=[2]" & IIf(lngCol = mlngOper, " And ��¼����=4", " And ��Ŀ���=[3]")
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ��¼+��ŵ����汾", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)), lng��Ŀ���)
    lng��ǰ�汾 = NVL(rsTemp!��ǰ�汾, 1)
    
    'ֻ�е�ǰ�汾������߰汾,���������(ǩ��������Ҳ���������)
    'ͬʱ�����߰汾=1,��ǩ����Ϊ��,Ҳ�������
    CheckVersion = ((lng��ǰ�汾 > lng��߰汾) Or (lng��߰汾 = 1 And vsf.Cell(flexcpForeColor, lngRow, lngCol) = &HFF&))
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    blnMsg = (strMsg <> "")
    gstrSQL = " Select ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(NVL(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����ID, lng��ҳID)
    With rsTemp
        .Filter = "����ID=" & mlng����ID
        Do While Not .EOF
            If strTime >= !��ʼʱ�� And strTime <= !��ֹʱ�� Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����ID, lng��ҳID, strTime, strCurTime) Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
        .Filter = "��ʼԭ��=1"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 1 And strTime < !��ʼʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And strTime < !��ʼʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And strTime > !��ֹʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '�������˵��
        strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]"
        GoTo exitHand
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Sub imgSign_Click()
    Call picSign_Click
End Sub

Private Sub lbl��֤ǩ��_Click()
    Call picSign_Click
End Sub

Private Sub picSign_Click()
    '����ǩ����ʷ��¼
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    vsfSignData.Clear
    str����ʱ�� = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    gstrSQL = "" & _
        " SELECT A.��¼�� AS ǩ����,NVL(to_char(A.�޸�ʱ��,'yyyy-MM-dd hh24:mi:ss'),A.��Ŀ����) AS ǩ��ʱ��,A.��¼���� AS ǩ����Ϣ,A.��¼��� AS ǩ������,A.ID,DECODE(A.��ĿID,NULL,'��Ч','δ��֤') AS ��Ч��,A.��ʼ�汾,A.��Ŀ��� AS ǩ������汾" & vbNewLine & _
        " FROM ���˻������� A,���˻����¼ B" & vbNewLine & _
        " WHERE A.��¼ID=B.ID AND A.��¼����=5" & vbNewLine & _
        " AND B.����ID=[1] AND B.��ҳID=[2] AND B.Ӥ��=[3] AND B.����ʱ��=[4] " & vbNewLine & _
        " Order by A.��Ŀ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����ʷ��¼", mlng����ID, mlng��ҳID, mintӤ��, CDate(str����ʱ��))
    
    Set vsfSignData.DataSource = rsTemp
    With vsfSignData
        .ColWidth(0) = 1000
        .ColWidth(1) = 1800
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .Row = 1
        .Col = 5
    End With
    
    picSign.Visible = False
    With picSignCheck
        .Left = vsf.Left + (vsf.Width - .Width) / 2
        .Top = vsf.Top + (vsf.Height - .Height) / 2
        .Visible = True
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    picSignCheck.Visible = False
End Sub

Private Sub cmdSignCur_Click()
    '������֤
    Dim lngLoop As Long
    Dim int�汾 As Integer
    Dim strSource As String, str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.Row, 4)) = 0) Then Exit Sub
    If (Val(vsfSignData.TextMatrix(vsfSignData.Row, 7)) < 2) Then
        MsgBox "����ǩ������仯���ϰ�ǩ�������ݲ�֧��ǩ��У�鹦�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    int�汾 = vsfSignData.TextMatrix(vsfSignData.Row, 6)
    str����ʱ�� = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    Set rsTemp = GetSignData(str����ʱ��, int�汾)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    Debug.Print "��֤ǩ����" & Now & vbCrLf & strSource
    
    '����ǩ��
    Err.Clear
    If gobjTendESign Is Nothing Then
        On Error Resume Next
        Set gobjTendESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjTendESign Is Nothing Then
            Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    If gobjTendESign Is Nothing Then
        MsgBox "����ǩ������δ����ȷ��װ����֤�������ܼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjTendESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.Row, 4)), 5) Then
        vsfSignData.TextMatrix(vsfSignData.Row, 5) = "��Ч"
        Call vsfSignData_EnterCell
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSignAll_Click()
    Dim lngSel As Long
    Dim lngRow As Long, lngRows As Long
    'ȫ����֤
    
    lngSel = vsfSignData.Row
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        vsfSignData.Row = lngRow
        Call cmdSignCur_Click
    Next
    vsfSignData.Row = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln�ⲿ As Boolean = False) As Boolean
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    '��ʾ��ʷǩ�����
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln�ⲿ Then
        If vsf.Col <> mlngSigner Then Exit Function
    End If
    If vsf.TextMatrix(vsf.Row, mlngSigner) = "" Then Exit Function
    
    str����ʱ�� = vsf.TextMatrix(vsf.Row, 1) & " " & vsf.TextMatrix(vsf.Row, 2) & ":00"
    gstrSQL = "" & _
        " SELECT A.��¼�� AS ǩ����,NVL(to_char(A.�޸�ʱ��,'yyyy-MM-dd hh24:mi:ss'),A.��Ŀ����) AS ǩ��ʱ��,A.��¼���� AS ǩ����Ϣ,A.��¼��� AS ǩ������,A.ID" & vbNewLine & _
        " FROM ���˻������� A,���˻����¼ B" & vbNewLine & _
        " WHERE A.��¼ID=B.ID AND A.��¼����=5" & vbNewLine & _
        " AND B.����ID=[1] AND B.��ҳID=[2] AND B.Ӥ��=[3] AND B.����ʱ��=[4] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����ʷ��¼", mlng����ID, mlng��ҳID, mintӤ��, CDate(str����ʱ��))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    With picSign
        .Top = vsf.Top + vsf.CellTop + vsf.CellHeight - .Height
        .Left = vsf.Left + vsf.CellLeft + 500
        .Visible = True
    End With
    ShowSignMarker = True
End Function

Private Sub vsfSignData_EnterCell()
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.Row, 5) <> "��Ч")
End Sub

Private Function GetSignData(ByVal str����ʱ�� As String, ByVal int�汾 As Integer) As ADODB.Recordset
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If int�汾 = 1 Then
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.�޸�ʱ��" & vbNewLine & _
            "  From ���˻������� a, ���˻����¼ b" & vbNewLine & _
            " Where b.����id = [1] And b.��ҳid = [2] And B.Ӥ��=[3] And b.����ʱ�� =[4]" & vbNewLine & _
            "   And a.��¼id = b.ID and A.��¼���� <>5 and A.��ʼ�汾=1" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    Else
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.�޸�ʱ��" & vbNewLine & _
            "  From ���˻������� a, ���˻����¼ b" & vbNewLine & _
            " Where b.����id = [1] And b.��ҳid = [2] And B.Ӥ��=[3] And b.����ʱ�� =[4]" & vbNewLine & _
            "   And a.��¼id = b.ID and A.��¼���� <>5" & vbNewLine & _
            "   and (A.��ʼ�汾=[5] or (A.��ʼ�汾 <[5] and A.��ֹ�汾 IS NULL) or (A.��ʼ�汾<[5] and A.��ֹ�汾>[5]))" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ���汾������", mlng����ID, mlng��ҳID, mintӤ��, CDate(str����ʱ��), int�汾)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '���ⲿ���������
    If Not ShowSignMarker(True) Then Exit Sub
    Call picSign_Click
End Sub
