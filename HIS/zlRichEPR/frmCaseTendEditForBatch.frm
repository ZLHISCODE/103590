VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendEditForBatch 
   Caption         =   "����¼��"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15615
   Icon            =   "frmCaseTendEditForBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15615
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvw����� 
      Height          =   1725
      Left            =   8520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3043
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgRow"
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
   Begin VB.PictureBox picQuery 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   3675
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5850
      Width           =   3675
      Begin zlRichEPR.usrTendEditor mfrmCaseTendEditForSinglePerson 
         Height          =   1875
         Left            =   420
         TabIndex        =   25
         Top             =   300
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3307
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   4950
      Top             =   2970
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
            Picture         =   "frmCaseTendEditForBatch.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendEditForBatch.frx":686E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLocate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   2115
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "�����Ŷ�λ"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   15
         TabIndex        =   15
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   0
      ScaleHeight     =   4485
      ScaleWidth      =   10545
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1350
      Width           =   10545
      Begin MSComctlLib.ListView lvwMultiSel 
         Height          =   1725
         Left            =   3090
         TabIndex        =   21
         Top             =   1350
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
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   945
         TabIndex        =   17
         Top             =   3810
         Visible         =   0   'False
         Width           =   945
         Begin VB.CommandButton cmdδ��˵�� 
            Caption         =   "�E"
            Height          =   225
            Left            =   630
            TabIndex        =   19
            Top             =   30
            Width           =   255
         End
         Begin VB.ComboBox cbo��λ 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txt���� 
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Vsf 
         Height          =   4335
         Left            =   0
         TabIndex        =   22
         Top             =   30
         Width           =   10425
         _cx             =   18389
         _cy             =   7646
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
         FixedCols       =   0
         RowHeightMin    =   600
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCaseTendEditForBatch.frx":D0D0
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
         Begin VB.ListBox lst���±�ʶ 
            Height          =   600
            ItemData        =   "frmCaseTendEditForBatch.frx":D132
            Left            =   900
            List            =   "frmCaseTendEditForBatch.frx":D134
            TabIndex        =   27
            Top             =   600
            Width           =   915
         End
      End
   End
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   15015
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   15015
      Begin VB.OptionButton optLevel 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   14160
         TabIndex        =   29
         Top             =   33
         Width           =   735
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   13320
         TabIndex        =   28
         Top             =   33
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.PictureBox PicPati 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   6240
         ScaleHeight     =   300
         ScaleWidth      =   1500
         TabIndex        =   7
         Top             =   0
         Width           =   1500
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   400
            Style           =   2  'Dropdown List
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   1065
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   60
            Width           =   360
         End
      End
      Begin VB.CommandButton cmd����� 
         Caption         =   "�����"
         Height          =   320
         Left            =   7740
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmdˢ�� 
         Caption         =   "ˢ��(&R)"
         Height          =   320
         Left            =   11220
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   915
      End
      Begin VB.ComboBox cbo����ȼ� 
         Height          =   300
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   450
         TabIndex        =   2
         Top             =   0
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   174063619
         UpDown          =   -1  'True
         CurrentDate     =   38702
      End
      Begin VB.Label lblEntry 
         AutoSize        =   -1  'True
         Caption         =   "¼�뷽ʽ"
         Height          =   180
         Left            =   12360
         TabIndex        =   30
         Top             =   65
         Width           =   720
      End
      Begin VB.Label lbl������嵥 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���з���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   8580
         TabIndex        =   11
         Top             =   60
         Width           =   2505
      End
      Begin VB.Label lbl�ȼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4380
         TabIndex        =   5
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   3
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   8340
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendEditForBatch.frx":D136
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22463
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
      Bindings        =   "frmCaseTendEditForBatch.frx":D9C8
      Left            =   450
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCaseTendEditForBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnData As Boolean                 '�����Ƿ���ȡ������?
Private mstrSel As String                   '������:1;����ĳ��Ԫ��:1.1
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mintPreDays As Long
Private mstrMaxDate As String
Private mstrSelItems As String              '�����û��������ӵ��У�����ˢ�º���������
Private mstrTime As String                  '�����ȡ���ݺ����Чʱ��,�Ա���ʱ���Ƿ�����ȡ���ݺ�������޸�
Private mstr����� As String
Private mlng������ As Long                  '��ǰ��ʾ�Ĳ�����
Private mbyt���� As Integer                     '-1 ����,0 ĸ�� ,1 Ӥ�� (�����ƿ���ѡ����������Ĭ��Ϊ����)
Private mblnRefresh As Boolean              '�Ƿ�ˢ�¹�����
Private mblnCheckVersion As Boolean
Private mobjExtendedBar As CommandBar
Private mstrScope As String
Private mdtOutbegin As Date, mdtOutEnd As Date

'���±�����ֵ,��ENTERCELL�и���
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mbyt����ȼ� As Long
Private mintӤ�� As Integer
Private mbln���� As Boolean                 '�Ƿ���Ҫ¼������
Private mstrPrivs As String

Private mlngOper As Long                    '�����к�
Private mlngSigner As Long                  'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngGroup As Long                   '���
Private mlngCert As Long                    '֤��ID

Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsPatient As New ADODB.Recordset           '��ǰ¼��Ļ����¼��Ŀ�嵥

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Private Const pסԺ��ʿվ As Long = 1262

Public Event AfterDataChanged()
Public Event AfterArchiveChanged()
Public Event AfterRefresh()
Public Event AfterSelChange(ByVal lngCert As Long)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'��¼�ϴ�ѡ����,����,�Ա�ˢ�º����¶�λ
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private Enum ������Ϣ
    ���� = 1
    ����ID
    ��ҳID
    �Ա�
    סԺ��
    ����
    ����ȼ�
    ���±�ʶ
    ��Ч������
End Enum

Private Sub cbo��λ_Click()
    If txt����.Enabled = False Or Val(cbo��λ.Tag) = 1 Then txt����.Text = cbo��λ.Text
End Sub

Private Sub cbo����ȼ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmd�����.SetFocus
End Sub

Private Sub cbo����_Click()
    Dim arrCode
    Dim blnVisible As Boolean
    Dim lngCount As Long
    
    If cbo����.ListCount = 0 Then Exit Sub
    If cbo����.Tag = "" Then cbo����.Tag = String(cbo����.ListCount, "[LPF]0"): cbo����.Tag = Mid(cbo����.Tag, 4)
    arrCode = Split(cbo����.Tag, "[LPF]")
    blnVisible = (Val(arrCode(cbo����.ListIndex)) = 1)
    PicPati.Enabled = blnVisible
    PicPati.Visible = blnVisible
    '������ؿؼ�����
    If blnVisible = False Then
        cmd�����.Left = PicPati.Left
    Else
        cmd�����.Left = PicPati.Left + PicPati.Width + 20
    End If
    lbl������嵥.Left = cmd�����.Left + cmd�����.Width + 10
    lvw�����.Left = lbl������嵥.Left
    cmdˢ��.Left = lbl������嵥.Left + lbl������嵥.Width + 50
    lblEntry.Left = cmdˢ��.Left + cmdˢ��.Width + 50
    optLevel(0).Left = lblEntry.Left + lblEntry.Width + 50
    optLevel(1).Left = optLevel(0).Left + optLevel(0).Width + 20
    picCond.Width = optLevel(1).Left + optLevel(1).Width + 10
    
    '�˵�
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar
    
    If mobjExtendedBar Is Nothing Then Exit Sub
    'ɾ�������˵���
    For lngCount = mobjExtendedBar.Controls.Count To 1 Step -1
        mobjExtendedBar.Controls(lngCount).Delete
    Next
    Set objExtendedBar = mobjExtendedBar
    With objExtendedBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picCond.hWnd
        cbrCustom.ToolTipText = "����"
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picLocate.hWnd
        cbrCustom.ToolTipText = "��λ"
    End With
End Sub

Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cbo����ȼ�.SetFocus
End Sub


'û�����ݵ�����°��¼�,��Ӧ����δ��˵��
'�̶�/��ʾ¼��������׾��������
'���������ݺ��¼��򵯳���λ��ʽ
'��*��С�����������ַ������¼�

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim blnData As Boolean      '��������Ϊ��
    Dim blnClear As Boolean     '�����?
    Dim strRecord As String
    Dim strSymbol As String
    Dim strSelItems As String
    Dim strDelItem As String
    Dim lngOrder As Long
    Dim introw As Integer, intCol As Integer, intRowSel As Integer, intColSel As Integer
    Dim intRow_ As Integer, intCol_ As Integer
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        mstrSel = mfrmCaseTendEditForSinglePerson.GetCopyData
    Case conMenu_Edit_PASTE
        Call PasteData
    Case conMenu_Edit_Clear
        '���ν������ݵ����ҳ���,����colData����Ϊ1,Ȼ����ѡ��Ԫ����������
        blnEnable = picInput.Visible
        introw = vsf.ROW
        intCol = vsf.Col
        intRowSel = vsf.RowSel
        intColSel = vsf.ColSel
        
        If vsf.ROW > vsf.RowSel Then introw = vsf.RowSel: intRowSel = vsf.ROW
        If vsf.Col > vsf.ColSel Then intCol = vsf.ColSel: intColSel = vsf.Col
        If intColSel >= mlngSigner Then intColSel = mlngSigner - 1
        
        For intRow_ = introw To intRowSel
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
        For introw = intRowSel To 1 Step -1
            blnData = False
            For intCol = IIf(Val(vsf.RowData(introw)) = 0, 1, ��Ч������) To intColSel
                If vsf.TextMatrix(introw, intCol) <> "" Then
                    blnData = True
                    Exit For
                End If
            Next
            If Not blnData Then
                If Val(vsf.TextMatrix(introw, mlngRecord)) <> 0 Then   '��ʷ��������
                    vsf.RowHidden(introw) = True
                Else
                    If introw <> vsf.Rows - 1 Then
                        vsf.RemoveItem introw               '�¼�¼ɾ��
                        '���ɾ���������²�����,ͬ��ɾ���ڲ���¼��(����еĻ���ճ��ʱ������)
                        mrsPatient.Filter = "��=" & introw
                        If mrsPatient.RecordCount <> 0 Then
                            mrsPatient.Delete
                            mrsPatient.Filter = ""
                            If mrsPatient.RecordCount <> 0 Then mrsPatient.MoveFirst
                            Do While Not mrsPatient.EOF
                                If mrsPatient!�� > introw Then
                                    mrsPatient!�� = mrsPatient!�� - 1
                                End If
                                mrsPatient.MoveNext
                            Loop
                            mrsPatient.UpdateBatch
                        End If
                    End If
                End If
            End If
        Next
        
        mrsPatient.Filter = 0
        mblnShow = False
        picInput.Visible = False
        
        '���ѡ������
        vsf.RowSel = vsf.ROW
        vsf.ColSel = vsf.Col
        If vsf.Enabled And vsf.Visible Then vsf.SetFocus
        If blnEnable Then Call Vsf_EnterCell
    Case conMenu_Edit_SPECIALCHAR
        strSymbol = frmInsSymbol.ShowMe(False, 0)
        Me.txt����.Text = Me.txt����.Text & strSymbol
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
        
        '���Ƕಡ��,���Դ˴�����ȼ�ʼ�մ�-1,Ӥ����-1
        strSelItems = frmTendItemChoose.ShowSelect(strSelItems, -1, -1, mlng����ID)
        If strSelItems = "" Then Exit Sub
        mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & strSelItems
        
        Call InsertColumn(strSelItems)
    Case conMenu_Edit_Delete
        '�����ѯ�б���������������ɾ��
        intCol = vsf.Col
        intRowSel = vsf.Rows - 1
        For introw = vsf.ROW To intRowSel
            If vsf.TextMatrix(introw, intCol) <> "" Or vsf.Cell(flexcpData, introw, intCol) <> 0 Then
                MsgBox "��ǰ��Ŀ�����ݣ�������ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
        
        Call DeleteColumn(intCol)
    Case conMenu_Edit_NewItem   '����
        '��������(��������)
        '��λ����������
        mrsPatient.Filter = "��=" & vsf.ROW
        If mrsPatient.RecordCount <> 0 Then
            strRecord = mrsPatient!����ID & "|" & mrsPatient!��ҳID & "|" & mrsPatient!����ID & "|" & mrsPatient!Ӥ�� & "|" & mrsPatient!����ȼ� & "|" & mrsPatient!����ȼ����� & "|" & NVL(mrsPatient!ƥ����)
        End If
        mrsPatient.Filter = 0
        
        '�������в����Ƶ�ǰ�еĲ��˻�����Ϣ
        introw = vsf.ROW + 1
        If Val(vsf.TextMatrix(introw - 1, ����ID)) = 0 Then Exit Sub
        
        vsf.Rows = vsf.Rows + 1
        vsf.RowPosition(vsf.Rows - 1) = introw
        'ͬһ�����˵Ķ�������,ֻ�е�һ�в���ʾ���˵���Ϣ
'        Vsf.TextMatrix(intRow, ����) = Vsf.TextMatrix(intRow - 1, ����)
'        Vsf.TextMatrix(intRow, �Ա�) = Vsf.TextMatrix(intRow - 1, �Ա�)
'        Vsf.TextMatrix(intRow, סԺ��) = Vsf.TextMatrix(intRow - 1, סԺ��)
'        Vsf.TextMatrix(intRow, ����) = Vsf.TextMatrix(intRow - 1, ����)
        vsf.TextMatrix(introw, ����ID) = vsf.TextMatrix(introw - 1, ����ID)
        vsf.TextMatrix(introw, ��ҳID) = vsf.TextMatrix(introw - 1, ��ҳID)
        vsf.Cell(flexcpAlignment, introw, 1, introw, ��Ч������ - 1) = flexAlignLeftCenter
        'Vsf.Cell(flexcpAlignment, intRow, ��Ч������, intRow, Vsf.Cols - 1) = flexAlignCenterCenter
        
        '�����ڴ��¼��
        With mrsPatient
            .Filter = "��>=" & introw
            Do While Not .EOF
                !�� = !�� + 1
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        '��ӵ�ǰ��
        strFields = "��|����ID|��ҳID|����ID|Ӥ��|����ȼ�|����ȼ�����|ƥ����"
        strValues = introw & "|" & strRecord
        Call Record_Add(mrsPatient, strFields, strValues)
        
    Case conMenu_Edit_Transf_Save '����
        If SaveME Then Call ShowMe(mfrmParent, mlng����ID, mstrPrivs, False, False)
    Case conMenu_Edit_Transf_Cancle 'ȡ��
        mstrSel = ""
        mblnShow = False
        picInput.Visible = False
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
        
        Call ReadData(True)
        mblnChange = False
        
        Call vsf_AfterRowColChange(1, 1, vsf.ROW, 1)
    Case conMenu_Tool_Sign          'ǩ��
        Call SignMe
    Case conMenu_Manage_ThingDel    'ȡ��ǩ��
        Call UnSignMe
    Case conMenu_Edit_ApplyTo       '��������
        vsf.Rows = vsf.Rows + 1
        If Not vsf.RowIsVisible(vsf.Rows - 1) Then
            vsf.TopRow = vsf.Rows - 1
            vsf.ROW = vsf.Rows - 1
        End If
        vsf.Col = ����
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit = False Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And mlng������ <> 0
    Case conMenu_Edit_PASTE
        Control.Enabled = (mstrSel <> "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")
    Case conMenu_Edit_SPECIALCHAR, conMenu_Edit_Append
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")
    Case conMenu_Edit_Clear 'ǩ�������ݲ��������
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And mblnCheckVersion
        
        '����Ƕ�ѡ,���������
        If vsf.RowSel <> vsf.ROW Or vsf.ColSel <> vsf.Col Then Control.Enabled = True
    Case conMenu_Edit_Delete
        Dim blnDel As Boolean
        If mrsSelItems.State = 1 Then
            mrsSelItems.Filter = "��=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                blnDel = (mrsSelItems!�̶� = 0)
            End If
            mrsSelItems.Filter = 0
        End If
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And blnDel
    Case conMenu_Edit_NewItem   '����
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")  'û�й鵵������
    Case conMenu_Edit_ApplyTo   '������
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")  'û�й鵵������
    Case conMenu_Edit_Transf_Save '����
        Control.Enabled = mblnChange Or (Format(dtp.Value, "yyyy-MM-dd HH:mm:ss") <> mstrTime And mblnData)
    Case conMenu_Edit_Transf_Cancle 'ȡ��
        Control.Enabled = mblnChange
    Case conMenu_Tool_Sign          'ǩ��
        Control.Enabled = Not mblnChange And (vsf.TextMatrix(vsf.ROW, mlngSigner) = "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) <> 0
    Case conMenu_Manage_ThingDel    'ȡ��ǩ��
        Control.Enabled = Not mblnChange And (vsf.TextMatrix(vsf.ROW, mlngSigner) <> "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) <> 0
    End Select
End Sub

Private Sub cmd�����_Click()
    Dim lvwItem As ListItem
    Dim intDo As Integer, intMax As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݵ�ǰѡ��Ŀ���,��ȡ������Ϣ
    
    gstrSQL = " Select distinct ����� From ��λ״����¼ Where ����ID=[1] " & IIf(CLng(cbo����.ItemData(Me.cbo����.ListIndex)) = -1, "", " And ����ID=[2]") & " And ����� is not null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, CLng(cbo����.ItemData(Me.cbo����.ListIndex)))
    With rsTemp
        lvw�����.ListItems.Clear
        lvw�����.ListItems.Add , "K0", "���з���", , 2
        
        Do While Not .EOF
            Set lvwItem = lvw�����.ListItems.Add(, "K" & .AbsolutePosition, !�����, , 2)
            '�����û���ѡ����ʾ
            If InStr(1, "," & lbl������嵥.Caption & ",", "," & !����� & ",") <> 0 Then lvwItem.Checked = True
            .MoveNext
        Loop
        If InStr(1, "," & lbl������嵥.Caption & ",", ",���з���,") <> 0 Then
            lvw�����.ListItems(1).Checked = True
            Call lvw�����_ItemCheck(lvw�����.ListItems(1))
        End If
        
        lvw�����.Move lbl������嵥.Left + 100, 900
        lvw�����.Visible = True
        lvw�����.SetFocus
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdˢ��_Click()
    If mblnInit = False Then Exit Sub
    
    mlng����ID = Me.cbo����.ItemData(Me.cbo����.ListIndex)
    mbyt����ȼ� = Me.cbo����ȼ�.ItemData(Me.cbo����ȼ�.ListIndex)
    mstr����� = Me.lbl������嵥.Caption
    If PicPati.Visible = True Then
        mbyt���� = cbo����.ItemData(cbo����.ListIndex)
    Else
        mbyt���� = -1
    End If
    mblnRefresh = True
    
    Call ReadData
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picMain.hWnd
    Case 2
        Item.Handle = picQuery.hWnd
    End Select
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = stbThis.Height
    
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsf.Visible And vsf.Enabled Then vsf.SetFocus
    End If
End Sub

Private Sub dtp_LostFocus()
    If Me.dtp.Tag = Me.dtp.Value Then Exit Sub
    Me.dtp.Tag = Me.dtp.Value
    
    If mblnRefresh Then
        '��ǰˢ�º������,���������ֵ
        mlng����ID = Me.cbo����.ItemData(Me.cbo����.ListIndex)
        mbyt����ȼ� = Me.cbo����ȼ�.ItemData(Me.cbo����ȼ�.ListIndex)
        mstr����� = Me.lbl������嵥.Caption
    End If
    
    Call ReadData
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

Private Sub Form_Activate()
    Call Vsf_EnterCell
    Call vsf_AfterRowColChange(1, 1, vsf.ROW, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If InStr(1, "TXT����,CBO��λ", UCase(Me.ActiveControl.Name)) <> 0 Then
            mblnShow = False
            picInput.Visible = False
            vsf.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    mstrSel = ""
    mstrSelItems = ""
    mstr����� = ""
    mblnShow = False
    mblnChange = False
    mblnInit = False
    mblnRefresh = False
    mlng����ID = 0
    mlng������ = 0
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    dtp.MaxDate = mstrMaxDate & " 23:59:59"
    dtp.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    Set mobjExtendedBar = Nothing
    
    Call InitMenuBar
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    Call InitPanelMain
    Call InitEnv            '��ʼ������
    
    lst���±�ʶ.Clear
    lst���±�ʶ.AddItem "1��/��"
    lst���±�ʶ.AddItem "2��/��"
    lst���±�ʶ.AddItem "3��/��"
    lst���±�ʶ.AddItem "4��/��"
    lst���±�ʶ.AddItem "5��/��"
    lst���±�ʶ.AddItem "6��/��"
End Sub

Private Sub InitEnv()
    Dim curDate As Date, intDay As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim blnVisible As Boolean
    Dim blnType As Boolean
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    mstrScope = zlDatabase.GetPara("������ʾ��Χ", glngSys, pסԺ��ʿվ, "10000")
    '��Ժ����ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, pסԺ��ʿվ, 7))
    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, pסԺ��ʿվ, 30))
    mdtOutbegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
    
    blnType = Val(GetSetting("ZLSOFT", "˽��ģ��\frmCaseTendEditForBatch\" & gstrUserName, "Value")) = 0
    If blnType Then
        optLevel(0).Value = True
    Else
        optLevel(1).Value = True
    End If
    
    '��ȡ��ǰ�����µ����п���
    gstrSQL = " Select distinct B.ID,B.����||'-'||B.���� AS ����,decode(nvl(E.��������,''),'����',1,0) ����" & _
              " From �������Ҷ�Ӧ A,���ű� B,������Ա C,��Ա�� D,��������˵�� E" & _
              " Where A.����ID = b.ID And A.����ID=C.����ID And C.��ԱID=D.ID And A.����ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "��ǰ����") <> 0, "", " And D.ID=[2]") & _
              " And B.ID=E.����ID(+) And E.��������(+)='����'" & _
              " Order by B.����||'-'||B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, glngUserId)
    With Me.cbo����
        .Clear
        .Tag = ""
        If InStr(1, mstrPrivs, "��ǰ����") <> 0 Then
            .AddItem "���п���"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!ID
            .Tag = .Tag & "[LPF]" & rsTemp!����
            If blnVisible = False Then blnVisible = (Val(rsTemp!����) = 1)
            rsTemp.MoveNext
        Loop
        .Tag = IIf(blnVisible = True, 1, 0) & .Tag
        If Left(.Tag, 5) = "[LPF]" Then .Tag = Mid(.Tag, 6)
        .ListIndex = 0
    End With
    
    '��ȡ���л���ȼ�
    With Me.cbo����ȼ�
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = -1
        .AddItem "��������"
        .ItemData(.NewIndex) = 3
        .AddItem "��������"
        .ItemData(.NewIndex) = 2
        .AddItem "һ������"
        .ItemData(.NewIndex) = 1
        .AddItem "�ؼ�����"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
    End With
    
    '��Ӳ���
    With Me.cbo����
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = -1
        .AddItem "ĸ��"
        .ItemData(.NewIndex) = 0
        .AddItem "Ӥ��"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Where B.Ӧ�÷�ʽ<>0 " & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
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
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane): objPane.Title = "��ѯ": objPane.Options = PaneNoCaption
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
        .SetIconSize True, 24, 24
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

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "���ӷ���(Alt+G)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "���"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�����Ŀ(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"):  cbrControl.ToolTipText = "ɾ����Ŀ(Alt+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "������(Ctrl+A)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(Alt+S)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��¼ǩ��(Alt+R)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ȡ��"):  cbrControl.ToolTipText = "ȡ��ǩ��(Alt+U)"
    End With
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    Set objExtendedBar = cbsThis.Add("����", xtpBarTop)
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    With objExtendedBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picCond.hWnd
        cbrCustom.ToolTipText = "����"
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picLocate.hWnd
        cbrCustom.ToolTipText = "��λ"
    End With
    
    Set mobjExtendedBar = objExtendedBar
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
        .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
        .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
        .Add FALT, Asc("C"), conMenu_Edit_Audit
        .Add FALT, Asc("N"), conMenu_Edit_NewItem
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
        .Add FALT, Asc("S"), conMenu_Edit_Transf_Save
        .Add FALT, Asc("R"), conMenu_Tool_Sign
        .Add FALT, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitBill()
    Dim blnLocate As Boolean            '�Ƿ��ҵ���ǰ��ʿ�����ܿ���
    Dim intCol As Integer, intCols As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ʼ���ڴ��¼��
    'mrsPatient
    strFields = "��," & adDouble & ",18|����ID," & adDouble & ",18|��ҳID," & adDouble & ",18|����ID," & adDouble & ",18|" & _
                "Ӥ��," & adDouble & ",18|����ȼ�," & adDouble & ",18|����ȼ�����," & adLongVarChar & ",50|ƥ����," & adLongVarChar & ",500"
    Call Record_Init(mrsPatient, strFields)
    'mrsSelItems
    strFields = "��," & adDouble & ",18|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2"
    Call Record_Init(mrsSelItems, strFields)
    strFields = "��|��Ŀ���|��Ŀ����|�̶�"
    
    '�����ģ���趨����Ŀ
    strSQL = " Select B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ,B.��Ŀ����,1 AS �̶�" & _
             " From ������Ŀģ�� A,�����¼��Ŀ B" & _
             " Where a.��Ŀ��� = b.��Ŀ��� And B.Ӧ�÷�ʽ<>0 And A.����ID=[1] And A.����ȼ�=-1 And B.���ò��� IN (0,1,2)" & _
             " Order by A.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.cbo����.ItemData(Me.cbo����.ListIndex))
    If rsTemp.RecordCount = 0 Then
        '����ǰ�Ĺ�����ȡ��Ŀ�嵥��¼��
        strSQL = " Select B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ,B.��Ŀ����,0 AS �̶�" & _
                 " From �����¼��Ŀ B" & _
                 " Where B.Ӧ�÷�ʽ<>0 And B.���ò��� IN (0,1,2)" & _
                 " Order by B.��Ŀ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    
    With vsf
        intCols = .Cols - 1
        For intCol = 1 To intCols
            .ColHidden(intCol) = False
        Next
        
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = rsTemp.RecordCount + .FixedCols + ��Ч������     '���������Ա𴲺�סԺ����,�ټ��Ϲ̶���������
        .RowHeightMin = 600
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .WordWrap = True
        
        .TextMatrix(0, ����) = "����"
        .TextMatrix(0, �Ա�) = "�Ա�"
        .TextMatrix(0, סԺ��) = "סԺ��"
        .TextMatrix(0, ����) = "����"
        .TextMatrix(0, ����ȼ�) = "����ȼ�"
        .TextMatrix(0, ����ID) = "����ID"
        .TextMatrix(0, ��ҳID) = "��ҳID"
        .TextMatrix(0, ���±�ʶ) = "���±�ʶ"
        .ColWidth(0) = 300
        .ColWidth(����) = 1700
        .ColWidth(�Ա�) = 500
        .ColWidth(סԺ��) = 1000
        .ColWidth(����) = 800
        .ColWidth(����ȼ�) = 1000
        .ColWidth(����ID) = 0
        .ColWidth(��ҳID) = 0
        .ColWidth(���±�ʶ) = 1000
        
        intCol = ��Ч������
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
            strFields = "��|��Ŀ���|��Ŀ����|�̶�"
            strValues = intCol & "|" & rsTemp!��Ŀ��� & "|" & rsTemp!��Ŀ���� & "|" & rsTemp!�̶�
            Call Record_Add(mrsSelItems, strFields, strValues)
            
            intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        '.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .MergeCells = flexMergeFree
        .WordWrap = True
        
        '��Ŀǰ��ѡ�����Ŀ�����ڴ��¼����
        strFields = "��|��Ŀ���|��Ŀ����|�̶�"
        strValues = .Cols - 1 & "|0|����|1"
        Call Record_Add(mrsSelItems, strFields, strValues)
        
        mlngOper = .Cols - 1
        .TextMatrix(0, .Cols - 1) = "����"
    End With
    
    '����Ƿ���Ҫ¼������
    mrsSelItems.Filter = "��Ŀ���=-1"
    mbln���� = (mrsSelItems.RecordCount <> 0)
    mrsSelItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AddPatient(ByVal StrKey As String, Optional ByVal rsPatient As ADODB.Recordset) As Boolean
    Dim lngRow As Long
    Dim strFind As String
    Dim strItems As String
    Dim strStart As String
    Dim intCol As Integer, intCols As Integer
    Dim int����ȼ� As Integer, intӤ�� As Integer, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If rsPatient Is Nothing Then
        strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
        '������,���ҵ��ò��˵���Ϣ,�������в����������(ֻ��ͨ������������ķ�������ͬһ���˵Ķ�������)
        lngRow = vsf.ROW
        StrKey = UCase(StrKey)
        
        Select Case Left(StrKey, 1)
        Case "-"    '����ID
            StrKey = Mid(StrKey, 2)
            strFind = " ����ID=" & StrKey
        Case "+"    'סԺ��
            StrKey = Mid(StrKey, 2)
            strFind = " סԺ��=" & StrKey
        Case Else   '����
            strFind = " ����='" & StrKey & "'"
        End Select
        '73204:������,2014-06-09,�޷�������Ժ��ƵĲ���
        '73097:������,2014-06-09,���סԺ�Ĳ��˻���¼���б���ֶ���(���a.סԺ����=b.��ҳID)
        '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
        '34�汾������Ϣ�����ҳID
        '��ȡ�����б�
        gstrSQL = " SELECT B.����ID, B.��ҳID, NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,G.��Ϣֵ AS ���±�ʶ, B.סԺ��, B.��Ժ���� AS ����,zl_PatitTendGrade(B.����ID,B.��ҳID) AS ����ȼ�,D.���� AS ����ȼ�����,B.��Ժ����ID AS ����ID,0 AS Ӥ��" & _
                  " FROM ������Ϣ A,������ҳ B,���ű� C,����ȼ� D,���˱䶯��¼ F,������ҳ�ӱ� G,��Ժ���� R" & _
                  " Where A.����ID = b.����ID And A.��ҳID=B.��ҳID And NVL(b.��ҳID, 0) <> 0 And b.��Ժ����ID = C.ID " & _
                  " And A.����ID=F.����ID And A.��ҳID=F.��ҳID And (F.��ʼԭ��=2 OR (F.��ʼԭ��=1 And Nvl(B.״̬,0)<>1 And NOT Exists(Select ����ID From ���˱䶯��¼ Where ��ʼԭ��=2 and ����ID=F.����ID And ��ҳID=F.��ҳID))) And F.��ʼʱ��<=[5]" & _
                  " And B.����ID=G.����ID(+) And B.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ' " & _
                  " AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL And B.����ȼ�ID=D.���(+) And R.����ID=A.����ID And R.����ID=[3] " & _
                  IIf(mlng����ID = -1, "", " And R.����ID=[4]")
        If Val(Mid(mstrScope, 2, 1)) <> 0 Then
            gstrSQL = gstrSQL & _
                  " Union" & _
                  " SELECT B.����ID, B.��ҳID, NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,G.��Ϣֵ AS ���±�ʶ, B.סԺ��, B.��Ժ���� AS ����,zl_PatitTendGrade(B.����ID,B.��ҳID) AS ����ȼ�,D.���� AS ����ȼ�����,B.��Ժ����ID AS ����ID,0 AS Ӥ��" & _
                  " FROM ������Ϣ A,������ҳ B,���ű� C,����ȼ� D,���˱䶯��¼ F,������ҳ�ӱ� G" & _
                  " Where A.����ID = b.����ID And A.��ҳID=B.��ҳID And NVL(b.��ҳID, 0) <> 0 And b.��Ժ����ID = C.ID And b.��ǰ����ID + 0 = [3]" & _
                  " And A.����ID=F.����ID And A.��ҳID=F.��ҳID And (F.��ʼԭ��=2 OR (F.��ʼԭ��=1 And NOT Exists(Select ����ID From ���˱䶯��¼ Where ��ʼԭ��=2 and ����ID=F.����ID And ��ҳID=F.��ҳID))) And F.��ʼʱ��<=[5]" & _
                  " And B.����ID=G.����ID(+) And B.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ' " & _
                  " AND B.��Ժ���� BETWEEN [1] AND [2] AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� is NULL And B.����ȼ�ID=D.���(+)" & _
                IIf(mlng����ID = -1, "", " And B.��Ժ����ID=[4]")
        End If
        '��ȡ�������б�
        gstrSQL = gstrSQL & _
                  " UNION " & _
                  " Select B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,G.��Ϣֵ AS ���±�ʶ,B.סԺ��,B.����,B.����ȼ�,B.����ȼ�����,B.����ID AS ����ID,A.��� AS Ӥ��" & _
                  " From ������������¼ A,(" & gstrSQL & ") B,������ҳ�ӱ� G" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
                  " And A.����ID=G.����ID(+) And A.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ'||DECODE(A.���,0,'',A.���) "
        
        gstrSQL = " Select * From (" & gstrSQL & ") " & _
                  " Where " & strFind
        Set rsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdtOutbegin, mdtOutEnd, mlng����ID, mlng����ID, CDate(strStart))
        If rsPatient.RecordCount = 0 Then
            MsgBox "û���ҵ��ò��ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        '�������,ֻ���ܷ����ڳ�ʼ����ʱ��
        lngRow = vsf.Rows - 1
    End If
    
    '��������ӵ������
    intCols = vsf.Cols - 1
    With rsPatient
        Do While Not .EOF
            If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
            
            vsf.TextMatrix(lngRow, ����) = IIf(!Ӥ�� <> 0, Space(4), "") & !����
            vsf.TextMatrix(lngRow, �Ա�) = !�Ա�
            vsf.TextMatrix(lngRow, סԺ��) = NVL(!סԺ��)
            vsf.TextMatrix(lngRow, ����) = NVL(!����)
            vsf.TextMatrix(lngRow, ����ȼ�) = NVL(!����ȼ�����)
            vsf.TextMatrix(lngRow, ����ID) = !����ID
            vsf.TextMatrix(lngRow, ��ҳID) = !��ҳID
            vsf.TextMatrix(lngRow, ���±�ʶ) = NVL(!���±�ʶ)
            
            '�����������(������ȷ��������Ϣǰ¼��������,��ȷ��������Ϣ,����һЩ����¼�����Ŀ��������
            vsf.Cell(flexcpData, lngRow, ��Ч������, lngRow, vsf.Cols - 1) = 0
            vsf.Cell(flexcpText, lngRow, ��Ч������, lngRow, vsf.Cols - 1) = ""
            
            If !����ȼ� <> int����ȼ� Or !Ӥ�� <> intӤ�� Or lng����ID <> !����ID Then
                strItems = ""
                int����ȼ� = !����ȼ�
                intӤ�� = !Ӥ��
                lng����ID = !����ID
                
                '��ȡ����������༭����Ŀ
                gstrSQL = " Select B.��Ŀ���" & _
                          " From �����¼��Ŀ B" & _
                          " Where B.Ӧ�÷�ʽ<>0 And B.����ȼ� >= [1] And B.���ò��� IN (0,[2])" & _
                          " And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[3])))"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(!����ȼ�), CInt(IIf(!Ӥ�� = 0, 1, 2)), CLng(!����ID))
                With rsTemp
                    Do While Not .EOF
                        strItems = strItems & IIf(strItems = "", "", ",") & !��Ŀ���
                        
                        .MoveNext
                    Loop
                End With
            End If
            
            '����Ϣ���µ��ڴ��¼����
            strFields = "��|����ID|��ҳID|����ID|Ӥ��|����ȼ�|����ȼ�����|ƥ����"
            strValues = lngRow & "|" & !����ID & "|" & !��ҳID & "|" & !����ID & "|" & !Ӥ�� & "|" & !����ȼ� & "|" & !����ȼ����� & "|" & strItems
            Call Record_Add(mrsPatient, strFields, strValues)
            
            Call DrawBackColor(lngRow)
            
            AddPatient = True
            lngRow = lngRow + 1
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    vsf.Cell(flexcpAlignment, 1, ����, vsf.Rows - 1, ��Ч������ - 1) = flexAlignLeftCenter
    If mblnInit Then Call vsf_AfterRowColChange(vsf.ROW + 1, 1, vsf.ROW, 1)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawBackColor(ByVal lngRow As Long)
    Dim intCol As Integer
    '��������¼�����Ŀ������Ϊ��ɫ
    
    mrsPatient.Filter = "��=" & lngRow
    If mrsPatient.RecordCount <> 0 Then
        For intCol = ��Ч������ To vsf.Cols - 1
            mrsSelItems.Filter = "��=" & intCol
            If mrsSelItems.RecordCount <> 0 Then
                If intCol <> mlngOper And InStr(1, "," & mrsPatient!ƥ���� & ",", "," & mrsSelItems!��Ŀ��� & ",") = 0 Then
                    vsf.Cell(flexcpBackColor, lngRow, intCol) = &HE0E0E0
                End If
            End If
        Next
    End If

    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
End Sub

Private Sub ReadData(Optional ByVal blnCancel As Boolean = False)
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    
    Dim int����Ӧ�� As Integer
    Dim strPatient As String, strChild As String
    Dim strStart As String, strEnd As String
    Dim rsData As New ADODB.Recordset
    Dim rsPatient As New ADODB.Recordset
    On Error GoTo errHand
    '��ȡ���ڶ����������
    
    If mblnChange And blnCancel = False Then
        If MsgBox("��ǰ���ݻ�δ���棬�㡰�ǡ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
    mblnShow = False
    picInput.Visible = False
    mblnInit = False
    
    mrsItems.Filter = "��Ŀ���=-1"
    If mrsItems.RecordCount <> 0 Then
        int����Ӧ�� = mrsItems!Ӧ�÷�ʽ
    End If
    mrsItems.Filter = 0
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    
    '1������ȡ����ѯʱ�䷶Χ���Լ���ӵ���Ŀ,���μӵ������
    Call InitBill
    Call AddColumns
    
    '���û�ѡ�������ӽ�ȥ
    If mstrSelItems <> "" Then
        arrColumn = Split(mstrSelItems, vbCrLf)
        intEnd = UBound(arrColumn)
        For intStart = 0 To intEnd
            Call InsertColumn(arrColumn(intStart))
        Next
    End If
    vsf.Cell(flexcpAlignment, 0, 0, 0, vsf.Cols - 1) = flexAlignCenterCenter
    '73204:������,2014-06-09,�޷�������Ժ��ƵĲ���
    '73097:������,2014-06-09,���סԺ�Ĳ��˻���¼���б���ֶ���(���a.סԺ����=b.��ҳID)
    '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
    '��ȡ�����б�
    strPatient = " SELECT B.����ID, B.��ҳID, NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,G.��Ϣֵ AS ���±�ʶ, B.סԺ��, B.��Ժ���� AS ����,zl_PatitTendGrade(B.����ID,B.��ҳID) AS ����ȼ�,D.���� AS ����ȼ�����,B.��Ժ����ID AS ����ID,0 AS Ӥ��" & _
              " FROM ������Ϣ A,������ҳ B,���ű� C,����ȼ� D,��λ״����¼ E,���˱䶯��¼ F,������ҳ�ӱ� G,��Ժ���� R" & _
              " Where A.����ID = B.����ID And A.��ҳID=B.��ҳID And NVL(b.��ҳID, 0) <> 0 And b.��Ժ����ID = C.ID " & _
              " AND Nvl(B.����״̬,0)<>5 AND B.���ʱ�� IS NULL And B.����ȼ�ID=D.���(+)" & _
              " And A.����ID=F.����ID And A.��ҳID=F.��ҳID And (F.��ʼԭ��=2 OR (F.��ʼԭ��=1 And Nvl(B.״̬,0)<>1 And NOT Exists(Select ����ID From ���˱䶯��¼ Where ��ʼԭ��=2 and ����ID=F.����ID And ��ҳID=F.��ҳID))) And F.��ʼʱ��<=[7]" & _
              " And B.����ID=G.����ID(+) And B.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ' " & _
              " AND A.����ID=E.����ID(+) And R.����ID=A.����ID And R.����ID=[3] " & IIf(mlng����ID = -1, "", " And R.����ID=[4]") & IIf(lbl������嵥.Caption = "���з���", "", " And instr([6],','||E.�����||',')<>0")
    '��ȡ�������б�
    strChild = " Select B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,G.��Ϣֵ AS ���±�ʶ,B.סԺ��,B.����,B.����ȼ�,B.����ȼ�����,B.����ID AS ����ID,A.��� AS Ӥ��" & _
              " From ������������¼ A,(" & strPatient & ") B,������ҳ�ӱ� G" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
              " And A.����ID=G.����ID(+) And A.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ'||DECODE(A.���,0,'',A.���) "
    If mbyt���� = 0 Then 'ĸ��
        strPatient = strPatient
    ElseIf mbyt���� = 1 Then 'Ӥ��
        strPatient = strChild
    Else
        strPatient = strPatient & " UNION " & strChild
    End If
'    strPatient = strPatient & _
'              " UNION " & _
'              " Select B.����ID,B.��ҳID,NVL(A.Ӥ������,B.����||'֮��'||A.���) AS ����,B.�Ա�,G.��Ϣֵ AS ���±�ʶ,B.סԺ��,B.����,B.����ȼ�,B.����ȼ�����,B.����ID AS ����ID,A.��� AS Ӥ��" & _
'              " From ������������¼ A,(" & strPatient & ") B,������ҳ�ӱ� G" & _
'              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
'              " And A.����ID=G.����ID(+) And A.��ҳID=G.��ҳID(+) And G.��Ϣ��(+)='���±�ʶ'||DECODE(A.���,0,'',A.���) "
    
    strPatient = "SELECT * FROM (" & strPatient & ") " & IIf(Me.cbo����ȼ�.ListIndex = 0, "", " WHERE ����ȼ�=[5]") & " Order by Lpad(����,10,' '),Ӥ��"
    Set rsPatient = zlDatabase.OpenSQLRecord(strPatient, Me.Caption, mdtOutbegin, mdtOutEnd, mlng����ID, mlng����ID, mbyt����ȼ�, "," & lbl������嵥.Caption & ",", CDate(strStart))
    Call AddPatient("", rsPatient)
    mlng������ = rsPatient.RecordCount
    
    '2����ȡ����
    gstrSQL = " Select X.* From ("
    If int����Ӧ�� = 2 Then
        gstrSQL = gstrSQL & _
                    "Select C.����ID,C.��ҳID,C.Ӥ��,A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                        "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,D.��Ŀ���� As ǩ��ʱ��," & _
                        "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,b.��¼���� As ���,b.��¼���," & _
                        "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,C.�鵵��,a.��¼�� " & _
                    " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                    " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                         " AND A.��¼���� =1 AND C.������Դ = 2 AND NVL(A.��¼���,0) <> 1 " & _
                         " And D.��¼����(+)=5 And D.��¼ID(+)=C.ID And D.��ֹ�汾(+) Is NULL" & _
                         " AND C.����ʱ�� = [7] "
    Else
        gstrSQL = gstrSQL & _
                    "Select C.����ID,C.��ҳID,C.Ӥ��,A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                        "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,D.��Ŀ���� As ǩ��ʱ��, " & _
                        "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,Decode(a.��Ŀ���,2,'',-1,'',b.��¼����) As ���,Decode(a.��Ŀ���,2,0,-1,0,b.��¼���) As ��¼���," & _
                        "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,C.�鵵��,a.��¼�� " & _
                    " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                    " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                         " AND A.��¼���� =1 AND C.������Դ = 2 AND ((NVL(A.��¼���,0) <> 1 And a.��Ŀ���>0) or a.��Ŀ���=-1 or (a.��Ŀ���=0 and a.��¼����=4)) " & _
                         " And D.��¼����(+)=5 And D.��¼ID(+)=C.ID And D.��ֹ�汾(+) Is NULL" & _
                         " AND C.����ʱ�� = [7] "
    End If
    gstrSQL = gstrSQL & _
                "       And a.��ֹ�汾 Is Null And b.��ֹ�汾 Is Null " & _
                "       And Decode(a.��Ŀ���,2,-1,a.��Ŀ���)=b.��Ŀ���(+)) X,�����¼��Ŀ Y " & _
                "Where Y.��Ŀ��� = X.��Ŀ��� And Nvl(y.Ӧ�÷�ʽ,0)=1 "
    
    '����������Ŀ
    gstrSQL = gstrSQL & _
                " UNION " & _
                " Select C.����ID,C.��ҳID,C.Ӥ��,A.��Ŀ���,DECODE(A.��¼����,4,A.��Ŀ����, A.��¼����) As ��¼���, " & _
                    "D.��ĿID AS ֤��ID,Nvl(A.��ֹ�汾,A.��ʼ�汾) AS ʵ�ʰ汾,D.��¼�� AS ǩ����,D.��Ŀ���� As ǩ��ʱ��, " & _
                    "Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,Decode(a.��Ŀ���,2,'',-1,'',b.��¼����) As ���,Decode(a.��Ŀ���,2,0,-1,0,b.��¼���) As ��¼���," & _
                    "C.����ʱ�� As �������,A.��¼id,A.��¼���,a.δ��˵��,C.�鵵��,a.��¼��" & _
                " From ���˻������� A, ���˻������� B,���˻����¼ C,���˻������� D " & _
                " Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                "      And a.��ֹ�汾 Is Null And b.��ֹ�汾 Is Null And D.��ֹ�汾(+) Is NULL" & _
                     " AND A.��¼���� =4 AND C.������Դ = 2 And D.��¼����(+)=5 And D.��¼ID(+)=C.ID AND C.����ʱ�� = [7] "
    
    gstrSQL = " Select A.*,B.����ID,B.����ȼ�,B.����,B.�Ա�,B.סԺ��,B.���� From (" & gstrSQL & ") A,(" & strPatient & ") B" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.Ӥ��=B.Ӥ��" & _
              " Order By A.����ID,A.��ҳID,A.Ӥ��,A.��¼���,DECODE(A.��Ŀ���,0,999,A.��Ŀ���)"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdtOutbegin, mdtOutEnd, mlng����ID, mlng����ID, mbyt����ȼ�, "," & lbl������嵥.Caption & ",", CDate(strStart))
    
    '׼���������(����û�е���Ŀ,ֱ���ڱ�������Ӹ���,ͬʱ�����ڲ���¼��
    Call ShowData(rsData)
    mstrTime = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    mblnData = (rsData.RecordCount)
    'Call OutputRsData(mrsSelItems)
    
    mblnInit = True
    If mlng������ <> 0 Then Call vsf_AfterRowColChange(2, 2, 1, 1)
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
    Dim lngRow As Long, lngRows As Long
    
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
    
    strFields = "��|��Ŀ���|��Ŀ����|�̶�"
    strValues = lngOrder & "|" & Split(strSelItems, "|")(0) & "|" & Split(strSelItems, "|")(1) & "|0"
    Call Record_Add(mrsSelItems, strFields, strValues)
    
    '���ģ������ĸ���
    mlngSigner = mlngSigner + 1
    mlngSignTime = mlngSignTime + 1
    mlngRecord = mlngRecord + 1
    mlngGroup = mlngGroup + 1
    mlngCert = mlngCert + 1
    
    '���ݲ������ô��еı���ɫ
    lngRows = vsf.Rows - 1
    For lngRow = 1 To lngRows
        Call DrawBackColor(lngRow)
    Next
End Sub

Private Sub AddColumns(Optional ByVal rsColumns As ADODB.Recordset)
    Dim blnAdd As Boolean
    '����ʷ�����д��ڵĶ�������ӵ������
    If Not rsColumns Is Nothing Then
        If rsColumns.State = 1 Then
            If rsColumns.RecordCount <> 0 Then
                blnAdd = True
            End If
        End If
    End If
    
    If blnAdd Then
        With rsColumns
            Do While Not .EOF
                mrsSelItems.Filter = "��Ŀ���=" & !��Ŀ���
                If mrsSelItems.RecordCount = 0 Then
                    mrsItems.Filter = "��Ŀ���=" & !��Ŀ���
                    vsf.Cols = vsf.Cols + 1
                    vsf.TextMatrix(0, vsf.Cols - 1) = .Fields("��Ŀ����").Value & IIf(NVL(mrsItems!��Ŀ��λ) = "", "", vbCrLf & "(" & mrsItems!��Ŀ��λ & ")")
                    vsf.ColAlignment(vsf.Cols - 1) = IIf(.Fields("��Ŀ����").Value = 0, flexAlignCenterCenter, flexAlignLeftTop)
                    mrsItems.Filter = 0
                    
                    strFields = "��|��Ŀ���|��Ŀ����|�̶�"
                    strValues = vsf.Cols - 1 & "|" & !��Ŀ��� & "|" & !��Ŀ���� & "|0"
                    Call Record_Add(mrsSelItems, strFields, strValues)
                End If
                .MoveNext
            Loop
        End With
    End If
    
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
        .TextMatrix(0, .Cols - 1) = "���"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngGroup = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "��¼��"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "�鵵��"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    mrsSelItems.Filter = 0
End Sub

Private Sub ShowData(ByVal rsData As ADODB.Recordset)
    On Error GoTo errHand
    Dim lngRow As Long
    Dim blnNewGroup As Boolean
    Dim lngGroup As Long        '���
    Dim lngRecord As Long       '��¼ID
    Dim strData As String       '�ݴ�����
    Dim strRecord As String     '��¼�ڴ��¼���е�����(��Ӧ��ȡ������Ϣ��)
    Dim str����ȼ����� As String
    Dim lng��ֹ�汾 As Long, bln��ɫ As Boolean
    Dim rsTemp As New ADODB.Recordset   '��ȡ��ǰ��¼������ֹ�汾
    
    strFields = "��|����ID|��ҳID|����ID|Ӥ��|����ȼ�|����ȼ�����|ƥ����"
    
    With rsData
        Do While Not .EOF
            '��¼ID��ͬ,��˵���ǲ�ͬ�Ĳ�����
            If (lngRecord <> !��¼ID Or lngGroup <> !��¼���) Then
                '��ȡ��ǰ��¼������ֹ�汾
                gstrSQL = " Select max(��ʼ�汾),Max(��ֹ�汾) From ���˻������� Where ��¼ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ��¼������ֹ�汾", CLng(!��¼ID))
                lng��ֹ�汾 = NVL(rsTemp.Fields(0).Value, 1)
                If lng��ֹ�汾 < NVL(rsTemp.Fields(1).Value, 1) Then lng��ֹ�汾 = NVL(rsTemp.Fields(1).Value, 1)
                
                blnNewGroup = False
                If lngRecord <> !��¼ID Then
                    '��λ����������
                    mrsPatient.Filter = "����ID=" & !����ID & " And ��ҳID=" & !��ҳID & " And ����ID=" & !����ID & " And Ӥ��=" & !Ӥ��
                    str����ȼ����� = "��������"
                    If mrsPatient.RecordCount <> 0 Then
                        mrsPatient.MoveLast
                        lngRow = mrsPatient!��
                        strRecord = mrsPatient!����ID & "|" & mrsPatient!��ҳID & "|" & mrsPatient!����ID & "|" & mrsPatient!Ӥ�� & "|" & mrsPatient!����ȼ� & "|" & mrsPatient!����ȼ����� & "|" & NVL(mrsPatient!ƥ����)
                        str����ȼ����� = mrsPatient!����ȼ�����
                    End If
                    mrsPatient.Filter = 0
                Else
                    '��������(��������)
                    blnNewGroup = True
                    lngRow = lngRow + 1
                    vsf.Rows = vsf.Rows + 1
                    vsf.RowPosition(vsf.Rows - 1) = lngRow
                    '�����ڴ��¼��
                    With mrsPatient
                        .Filter = "��>=" & lngRow
                        Do While Not .EOF
                            !�� = !�� + 1
                            .Update
                            .MoveNext
                        Loop
                        .Filter = 0
                    End With
                    '��ӵ�ǰ��
                    strValues = lngRow & "|" & strRecord
                    Call Record_Add(mrsPatient, strFields, strValues)
                    
                    Call DrawBackColor(lngRow)
                End If
                
                '��д��ǩ���˼�ǩ��ʱ��
                lngRecord = !��¼ID
                lngGroup = !��¼���
                bln��ɫ = True
                If Not IsNull(!ǩ����) Then
                    bln��ɫ = False
                    vsf.Cell(flexcpPicture, lngRow, 0) = imgRow.ListImages(1).Picture
                End If
                vsf.Cell(flexcpPictureAlignment, lngRow, 0) = flexAlignCenterCenter
                
'                If Not blnNewGroup Then
                    vsf.TextMatrix(lngRow, ����) = IIf(!Ӥ�� <> 0, Space(4), "") & CStr(.Fields("����").Value)
                    vsf.TextMatrix(lngRow, �Ա�) = CStr(.Fields("�Ա�").Value)
                    If NVL(.Fields("סԺ��").Value, 0) = 0 Then
                        vsf.TextMatrix(lngRow, סԺ��) = ""
                    Else
                        vsf.TextMatrix(lngRow, סԺ��) = CLng(NVL(.Fields("סԺ��").Value, 0))
                    End If
                    vsf.TextMatrix(lngRow, ����) = NVL(.Fields("����").Value)
                    vsf.TextMatrix(lngRow, ����ȼ�) = str����ȼ�����
'                End If
                vsf.TextMatrix(lngRow, ����ID) = Val(.Fields("����ID").Value)
                vsf.TextMatrix(lngRow, ��ҳID) = Val(.Fields("��ҳID").Value)
                vsf.TextMatrix(lngRow, mlngCert) = Val(NVL(.Fields("֤��ID").Value, 0))
                vsf.TextMatrix(lngRow, mlngSigner) = NVL(.Fields("ǩ����").Value)
                vsf.TextMatrix(lngRow, mlngSignTime) = Format(.Fields("ǩ��ʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                vsf.TextMatrix(lngRow, mlngRecord) = CLng(.Fields("��¼ID").Value)
                vsf.TextMatrix(lngRow, mlngGroup) = CLng(.Fields("��¼���").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 2) = NVL(.Fields("��¼��").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 1) = NVL(.Fields("�鵵��").Value)
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
    
    'ʹ��CellData�������޸ı�־
    vsf.Cell(flexcpAlignment, 1, 1, vsf.Rows - 1, ��Ч������ - 1) = flexAlignLeftCenter
    'Vsf.Cell(flexcpAlignment, 1, ��Ч������, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
    vsf.Cell(flexcpData, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Sub

Private Sub PasteData()
    Dim arrOrder, arrData
    Dim intCol As Integer, intCols As Integer
    '������Ŀ��Ž��и���
    
    arrOrder = Split(mstrSel, "|")
    arrData = Split(arrOrder(1), ",")
    arrOrder = Split(arrOrder(0), ",")
    intCols = UBound(arrOrder)
    
    mrsPatient.Filter = "��=" & vsf.ROW
    If mrsPatient.RecordCount <> 0 Then
        For intCol = 0 To intCols
            '���ȼ������Ƿ�ƥ��
            If InStr(1, "," & mrsPatient!ƥ���� & ",", "," & arrOrder(intCol) & ",") <> 0 Then
                '��������
                mrsSelItems.Filter = "��Ŀ���=" & arrOrder(intCol)
                If mrsSelItems.RecordCount <> 0 Then
                    If arrData(intCol) <> "" Or vsf.TextMatrix(vsf.ROW, mrsSelItems!��) <> "" Then
                        vsf.TextMatrix(vsf.ROW, mrsSelItems!��) = arrData(intCol)
                        
                        '�����޸ı�־
                        vsf.RowData(vsf.ROW) = 1
                        vsf.Cell(flexcpData, vsf.ROW, mrsSelItems!��) = 1
                        mblnChange = True   '�п��ܸ�����Ŀ�Ļ���ȼ������ò��˵�ǰ����ȼ��������ѭ��������
                    End If
                End If
            End If
        Next
    End If
    
    mrsPatient.Filter = 0
    mrsSelItems.Filter = 0
End Sub

Private Function WriteIntoVsf(Optional ByVal strInfo As String) As Boolean
    Dim blnAllow As Boolean
    Dim StrText As String
    Dim lngRecord As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long, strֵ�� As String
    
    lngRow = Split(txt����.Tag, "|")(0)
    lngCol = Split(txt����.Tag, "|")(1)
    
    If picInput.Visible Then
        If lngCol = ���� Then
            '¼�벡����Ϣ
            If AddPatient(txt����.Text) = False Then
                '��ԭ
                vsf.TextMatrix(lngRow, lngCol) = picInput.Tag
                txt����.Tag = ""
                picInput.Visible = False
            End If
        ElseIf txt����.Enabled Then
            '������ݺϷ���
             '������ݺϷ���
            If Val(cbo��λ.Tag) = 0 Then
                If txt����.Text <> "" Then
                    StrText = IIf(cbo��λ.Visible And Trim(cbo��λ.Text) <> "", cbo��λ.Text & ":", "") & Trim(txt����.Text)
                End If
            Else
                StrText = IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), cbo��λ.Text)
            End If
    
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
            
            blnAllow = CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, strֵ��)
            mrsItems.Filter = 0
            mrsSelItems.Filter = 0
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

Private Sub lst���±�ʶ_DblClick()
    Dim lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer
    On Error GoTo errHand
    '���没�����±�ʶ
    mrsPatient.Filter = "��=" & vsf.ROW
    If mrsPatient.RecordCount <> 0 Then
        '�ȶ�λ�޸Ĺ�����,��������ѭ���ҵ��޸Ĺ�����
        lng����ID = mrsPatient!����ID
        lng��ҳID = mrsPatient!��ҳID
        intӤ�� = mrsPatient!Ӥ��
        
        gstrSQL = "ZL_�������±�ʶ_Update(" & lng����ID & "," & lng��ҳID & "," & _
            intӤ�� & ",'" & lst���±�ʶ.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�����±�ʶ")
        vsf.TextMatrix(vsf.ROW, ���±�ʶ) = lst���±�ʶ.Text
        
        Call vsf_KeyDown(vbKeyReturn, 0)
    End If
    mrsPatient.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsPatient.Filter = 0
End Sub

Private Sub lst���±�ʶ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lst���±�ʶ_DblClick
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange Then
        If MsgBox("��ǰ���ݻ�δ���棬�㡰�ǡ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
End Sub

Private Sub lvw�����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Static blnExit As Boolean
    Dim blnFind As Boolean
    Dim blnCheck As Boolean
    Dim intDo As Integer, intMax As Integer
    
    If blnExit Then Exit Sub
    intMax = lvw�����.ListItems.Count
    blnCheck = Item.Checked
    
    If Item.Text = "���з���" Then
        blnExit = True
        For intDo = 2 To intMax
            lvw�����.ListItems(intDo).Checked = blnCheck
        Next
    Else
        blnExit = True
        For intDo = 2 To intMax
            If lvw�����.ListItems(intDo).Checked = Not blnCheck Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind Then
            lvw�����.ListItems(1).Checked = False
        Else
            lvw�����.ListItems(1).Checked = blnCheck
        End If
    End If
    
    blnExit = False
End Sub

Private Sub lvw�����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call lvw�����_Validate(blnCancel)
    cmdˢ��.SetFocus
End Sub

Private Sub lvw�����_Validate(Cancel As Boolean)
    Dim str�嵥 As String
    Dim intDo As Integer, intMax As Integer
    
    If lvw�����.ListItems(1).Checked Then
        lvw�����.Visible = False
        Me.lbl������嵥.Caption = "���з���"
        Exit Sub
    End If
    
    intMax = lvw�����.ListItems.Count
    For intDo = 2 To intMax
        If lvw�����.ListItems(intDo).Checked Then str�嵥 = str�嵥 & IIf(str�嵥 = "", "", ",") & lvw�����.ListItems(intDo).Text
    Next
    
    If str�嵥 = "" Then
        Cancel = True
        MsgBox "����ѡ��һ�����䣡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lvw�����.Visible = False
    lbl������嵥.Caption = str�嵥
    cmdˢ��.SetFocus
End Sub

Private Sub mfrmCaseTendEditForSinglePerson_DBCLICK(ByVal strData As String)
    Dim StrText As String
    
    If vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) <> "" Then Exit Sub
    
    mrsSelItems.Filter = "��Ŀ���=" & Val(Split(strData, "|")(0))
    If mrsSelItems.RecordCount <> 0 Then
        StrText = vsf.TextMatrix(vsf.ROW, mrsSelItems!��)
        
        '�����޸ı�־
        If StrText <> Split(strData, "|")(1) Then
            '�޸ı�־
            vsf.RowData(vsf.ROW) = 1
            vsf.Cell(flexcpData, vsf.ROW, mrsSelItems!��) = 1
            vsf.TextMatrix(vsf.ROW, mrsSelItems!��) = Split(strData, "|")(1)
            
            mblnChange = True
        End If
        
'        '�ƶ�����һ��
'        Call vsf_KeyDown(vbKeyReturn, 0)
    End If
    mrsSelItems.Filter = 0
End Sub


Private Sub optLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If optLevel(0).Value Then
         SaveSetting "ZLSOFT", "˽��ģ��\frmCaseTendEditForBatch\" & gstrUserName, "Value", 0
    Else
        SaveSetting "ZLSOFT", "˽��ģ��\frmCaseTendEditForBatch\" & gstrUserName, "Value", 1
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

Private Sub picQuery_Resize()
    mfrmCaseTendEditForSinglePerson.Left = 0
    mfrmCaseTendEditForSinglePerson.Top = 0
    mfrmCaseTendEditForSinglePerson.Width = picQuery.Width
    mfrmCaseTendEditForSinglePerson.Height = picQuery.Height
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    '��λ����ǰ������
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt����.Text) = "" Then Exit Sub
    
    lngRow = vsf.FindRow(Trim(txt����.Text), , ����)
    If lngRow < 1 Then Exit Sub
    vsf.ROW = lngRow
    If vsf.RowIsVisible(lngRow) = False Then vsf.TopRow = lngRow
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrText As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
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
                Me.cbo��λ.Clear
                Do While Not .EOF
                    Me.cbo��λ.AddItem !����
                    .MoveNext
                Loop
                Me.cbo��λ.ListIndex = 0
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
                strData = Trim(txt����.Text)
            End If
            If strData <> picInput.Tag Then blnAllow = WriteIntoVsf
        End If
        
        If blnAllow Then
            Call vsf_KeyDown(vbKeyReturn, Shift)
        Else
            Call Vsf_EnterCell
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If txt����.SelStart = 0 Then Call vsf_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        Dim lng����ID As Long, lng��ҳID As Long
        If Not (cmdδ��˵��.Visible And cbo��λ.Visible = False) Then Exit Sub
        
        mrsPatient.Filter = "��=" & vsf.ROW
        If mrsPatient.RecordCount <> 0 Then
            lng����ID = mrsPatient!����ID
            lng��ҳID = mrsPatient!��ҳID
        End If
        mrsPatient.Filter = 0
        
        If lng����ID = 0 Then Exit Sub
        StrText = frmWordsEditor.ShowMe(Me, lng����ID, lng��ҳID, txt����.Text)
        If StrText = "" Then Exit Sub
        txt����.Text = StrText
        
        Call txt����_KeyDown(vbKeyReturn, 0)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDo As Boolean
    Dim lng֤��ID As Long, strǩ���� As String, strǩ��ʱ�� As String
    Dim lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer
    '��ʾָ�����˵���ʷ����
    
    If mblnInit = False Then Exit Sub
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    stbThis.Panels(2).Text = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    stbThis.Panels(2).Text = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                stbThis.Panels(2).Text = ""
            End If
            
            If mrsSelItems!��Ŀ��� = 1 Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "�����±�ʾ��:39/37.5"
            ElseIf mrsSelItems!��Ŀ��� = 3 Then
                If mbln���� = False Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "������׾��ʾ��:130/120"
            ElseIf vsf.TextMatrix(0, NewCol) Like "Ѫѹ*" Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "¼�����:����ѹ/����ѹ"
            End If
            
            If mrsSelItems!��Ŀ��� >= 1 And mrsSelItems!��Ŀ��� <= 3 Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "�������в�λѡ��;��SHIFT+������δ��˵����ѡ��"
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '----------------------------------------
    '���ѡ���з����仯����ȡ�ò��˵���ʷ����
    If OldRow = NewRow Then Exit Sub
    '�������û�з����仯,Ҳ�����κδ���
    mrsPatient.Filter = "��=" & OldRow
    If mrsPatient.RecordCount <> 0 Then
        lng����ID = mrsPatient!����ID
        lng��ҳID = mrsPatient!��ҳID
        intӤ�� = mrsPatient!Ӥ��
    End If
    
    mrsPatient.Filter = "��=" & NewRow
    If mrsPatient.RecordCount <> 0 Then
        If lng����ID <> mrsPatient!����ID Or lng��ҳID <> mrsPatient!��ҳID Or intӤ�� <> mrsPatient!Ӥ�� Then
            blnDo = True
        End If
    End If
    
    '���߲�ѯ�Ӵ����������
    If blnDo Then
        Call mfrmCaseTendEditForSinglePerson.ShowMe(Me, mrsPatient!����ID, mrsPatient!��ҳID, mrsPatient!����ID, mrsPatient!Ӥ��, mrsPatient!����ȼ�, "", False, False)
        
        '���ݹ鵵���,ʵʱ���¹鵵��
        mrsPatient.Filter = "����ID=" & mrsPatient!����ID & " And ��ҳID=" & mrsPatient!��ҳID & " And Ӥ��=" & mrsPatient!Ӥ��
        Do While Not mrsPatient.EOF
            vsf.TextMatrix(mrsPatient!��, vsf.Cols - 1) = mfrmCaseTendEditForSinglePerson.mstrPigeonhole
            mrsPatient.MoveNext
        Loop
    End If
    
    mrsPatient.Filter = 0
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    blnScroll = True
    Call Vsf_EnterCell
    blnScroll = False
End Sub

Private Sub vsf_AfterUserResize(ByVal ROW As Long, ByVal Col As Long)
    Call Vsf_EnterCell
End Sub

Private Sub vsf_DblClick()
    mblnShow = True
    Call Vsf_EnterCell
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
            strData = txt����.Text
        End If
        If strData <> picInput.Tag Then blnAllow = WriteIntoVsf
    ElseIf lvwMultiSel.Visible Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
    ElseIf vsf.Col = ���±�ʶ Then
        blnAllow = True
    End If
    Call vsf.AutoSize(0, vsf.Cols - 1)
    picInput.Visible = False
    lvwMultiSel.Visible = False
    lst���±�ʶ.Visible = False
    If blnAllow = False Then
        If vsf.ROW <> lngRow Then vsf.ROW = lngRow
        If vsf.Col <> lngCol Then vsf.Col = lngCol
        Exit Sub
    End If
    
    RaiseEvent AfterSelChange(IIf(Trim(vsf.TextMatrix(vsf.ROW, mlngSigner)) <> "", 1, 0))
    
    mblnCheckVersion = CheckVersion
    If mblnShow = False Or (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) <> "") Then Exit Sub
    If vsf.Col = 0 Or vsf.ROW = 0 Then Exit Sub
    If vsf.Col = mlngOper And mblnCheckVersion = False Then Exit Sub
    If vsf.Col <> ���±�ʶ Then
        If vsf.Col >= mlngSigner Or (vsf.Col < ��Ч������ And (vsf.Col <> ���� And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) = 0)) Then Exit Sub   'ǩ����,ǩ��ʱ���Լ���Ų�����༭,�������
    End If
    If vsf.RowIsVisible(vsf.ROW) = False Then Exit Sub
    '�������Ƿ�����༭(���ڸò���,�ҵ���ǰ�е���Ŀ��Ž��м��)
    blnAllow = False
    If vsf.Col <> 1 Then
        If vsf.Col = ���±�ʶ Then
            With lst���±�ʶ
                .Left = vsf.CellLeft
                .Top = vsf.CellTop
                .Width = vsf.CellWidth
                .Visible = True
                .ZOrder 0
            End With
        Else
            mrsPatient.Filter = "��=" & vsf.ROW
            If mrsPatient.RecordCount <> 0 Then
                If vsf.Col = ���� Or vsf.Col = mlngOper Then
                    blnAllow = True
                Else
                    mrsSelItems.Filter = "��=" & vsf.Col
                    If mrsSelItems.RecordCount <> 0 Then
                        If InStr(1, "," & NVL(mrsPatient!ƥ����) & ",", "," & mrsSelItems!��Ŀ��� & ",") <> 0 Then
                            blnAllow = True
                        End If
                    End If
                End If
            End If
            mrsPatient.Filter = 0
            mrsSelItems.Filter = 0
        End If
    Else
        '�����Ŀհ���,ֻ����༭����
        blnAllow = True
    End If
    If Not blnAllow Then Exit Sub
    If Not blnScroll And vsf.Visible And vsf.Enabled Then vsf.SetFocus
    
    '׼����ʾ
    With picInput
        .Tag = vsf.TextMatrix(vsf.ROW, vsf.Col)             '����༭ǰ������
        .Left = vsf.ColPos(vsf.Col) + vsf.Left
        .Top = vsf.RowPos(vsf.ROW) + vsf.Top
        .Width = vsf.ColWidth(vsf.Col)
        If vsf.ROW = vsf.Rows - 1 Then
            .Height = vsf.ROWHEIGHT(vsf.ROW)    'ȡ���и�
        Else
            .Height = vsf.RowPos(vsf.ROW + 1) - vsf.RowPos(vsf.ROW)
        End If
        If .Height > vsf.RowHeightMax Then .Height = vsf.RowHeightMax
        If .Height < 600 Then .Height = 600
        .ZOrder 0
        .Visible = True
    End With
    With cbo��λ
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
                            If InStr(1, "," & vsf.TextMatrix(vsf.ROW, vsf.Col) & ",", "," & strData & ",") <> 0 Then lvwMultiSel.ListItems(intCol + 1).Checked = True
                        Next
                        lvwMultiSel.Tag = vsf.ROW & "|" & vsf.Col
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
        .Enabled = blnAllow          '�����ǰ���������л�ѡ��,������¼��
        .Text = vsf.TextMatrix(vsf.ROW, vsf.Col)
        If .Enabled Then
            If InStr(1, .Text, ":") <> 0 And cbo��λ.ListCount > 0 Then
                With cbo��λ
                    If InStr(1, txt����.Text, ":") <> 0 Then
                        Call zlControl.CboLocate(cbo��λ, Split(txt����.Text, ":")(0))
                    End If
                    '.Top = picInput.Height - .Height
                    .Width = picInput.Width
                    .Visible = True
                    .ZOrder 0
                End With
                .Text = Split(.Text, ":")(1)
            End If
        Else
            If .Text <> "" Then Call zlControl.CboLocate(cbo��λ, .Text)
            With cbo��λ
                '.Top = picInput.Height - .Height
                .Width = picInput.Width
                .Visible = True
                .ZOrder 0
            End With
        End If
        .Width = picInput.Width
        .Height = picInput.Height - IIf(cbo��λ.Visible, cbo��λ.Height, 0)
        .Tag = vsf.ROW & "|" & vsf.Col
    End With
    
    If cbo��λ.Enabled Then
        cbo��λ.Top = picInput.Height - cbo��λ.Height
        cbo��λ.Width = txt����.Width
    End If
    
    cmdδ��˵��.Visible = (InStr(1, "������������", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0) Or blnWords
    If cmdδ��˵��.Visible Then
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
    Dim strCompare As String
    Dim blnNextRow As Boolean
    
    '�������������,�Ե�
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 _
        Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        Exit Sub
    End If
    If KeyCode = vbKeyLeft And (picInput.Visible = False And lvwMultiSel.Visible = False) Then Exit Sub
    
    blnNextRow = Val(GetSetting("ZLSOFT", "˽��ģ��\frmCaseTendEditForBatch\" & gstrUserName, "Value")) = 0
    
    If KeyCode = vbKeyDelete Then
        '�����ǰ��Ԫ�������
        vsf.TextMatrix(vsf.ROW, vsf.Col) = ""
        Me.cbo��λ.Visible = False
        Me.txt����.Text = ""
        Me.txt����.Height = picInput.Height
    End If
    
    If KeyCode = vbKeyReturn Then
        '�����:56592,����,����¼��������ת.
        If blnNextRow = False Then
toNextRow2:
            If vsf.ROW < vsf.Rows - 1 Then
                vsf.ROW = vsf.ROW + 1
                vsf.Cell(flexcpAlignment, 1, ����, vsf.Rows - 1, ��Ч������ - 1) = flexAlignLeftCenter
                If vsf.RowHidden(vsf.ROW) Then GoTo toNextRow2
            Else
toNextCol2:
                If vsf.Col < mlngSigner Then
                    vsf.ROW = 1
                    vsf.Col = vsf.Col + 1
                    If vsf.Col = mlngSigner Then GoTo toNextCol2
                    If vsf.ColHidden(vsf.Col) Or vsf.Col < ��Ч������ Then GoTo toNextCol2
                    
                    '������¼�����ֱ������
                    If vsf.Col <> mlngOper Then
                        mrsPatient.Filter = "��=" & vsf.ROW
                        If mrsPatient.RecordCount <> 0 Then
                            strCompare = mrsPatient!ƥ����
                            mrsSelItems.Filter = "��=" & vsf.Col
                            If strCompare <> "" Then    'Ϊ����˵���������¼���,��û��¼�벡����Ϣ
                                If InStr(1, "," & strCompare & ",", "," & mrsSelItems!��Ŀ��� & ",") = 0 Then GoTo toNextCol2
                            End If
                        End If
                        mrsPatient.Filter = 0
                        mrsSelItems.Filter = 0
                    End If
                Else
                    vsf.ROW = 1
                    vsf.Col = mlngSigner - 1
                End If
            End If

            If vsf.ColIsVisible(vsf.Col) = False Then
                vsf.LeftCol = vsf.Col
            End If
            If vsf.RowIsVisible(vsf.ROW) = False Then
                vsf.TopRow = vsf.ROW
            End If
            Exit Sub

        
        Else
        
    
        '������һ����Ч��Ԫ��
toNextCol:
            If vsf.Col < mlngSigner Then
                vsf.Col = vsf.Col + 1
                If vsf.Col = mlngSigner Then GoTo toNextCol
                If vsf.ColHidden(vsf.Col) Or vsf.Col < ��Ч������ Then GoTo toNextCol
                
                '������¼�����ֱ������
                If vsf.Col <> mlngOper Then
                    mrsPatient.Filter = "��=" & vsf.ROW
                    If mrsPatient.RecordCount <> 0 Then
                        strCompare = mrsPatient!ƥ����
                        mrsSelItems.Filter = "��=" & vsf.Col
                        If strCompare <> "" Then    'Ϊ����˵���������¼���,��û��¼�벡����Ϣ
                            If InStr(1, "," & strCompare & ",", "," & mrsSelItems!��Ŀ��� & ",") = 0 Then GoTo toNextCol
                        End If
                    End If
                    mrsPatient.Filter = 0
                    mrsSelItems.Filter = 0
                End If
            Else
toNextRow:
                If vsf.ROW = vsf.Rows - 1 Then
                    vsf.Rows = vsf.Rows + 1
                    vsf.Cell(flexcpAlignment, 1, ����, vsf.Rows - 1, ��Ч������ - 1) = flexAlignLeftCenter
                    'Vsf.Cell(flexcpAlignment, Vsf.Rows - 1, ��Ч������, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
                End If
                vsf.ROW = vsf.ROW + 1
                If vsf.RowHidden(vsf.ROW) Then GoTo toNextRow
                vsf.Col = 1
            End If
            If vsf.ColIsVisible(vsf.Col) = False Then
                vsf.LeftCol = vsf.Col
            End If
            If vsf.RowIsVisible(vsf.ROW) = False Then
                vsf.TopRow = vsf.ROW
            End If
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        '������һ����Ч��Ԫ��
toPreCol:
        If vsf.Col > 1 Then
            vsf.Col = vsf.Col - 1
            If vsf.Col >= mlngSigner Then GoTo toPreCol
            If vsf.Col = mlngOper Then GoTo toPreCol
            If vsf.Col <> 1 And vsf.Col < ��Ч������ Then GoTo toPreCol
            If vsf.ColHidden(vsf.Col) Then GoTo toPreCol
            
            '������¼�����ֱ������
            If vsf.Col <> 1 Then
                mrsPatient.Filter = "��=" & vsf.ROW
                If mrsPatient.RecordCount <> 0 Then
                    strCompare = mrsPatient!ƥ����
                    mrsSelItems.Filter = "��=" & vsf.Col
                    If strCompare <> "" Then    'Ϊ����˵���������¼���,��û��¼�벡����Ϣ
                        If InStr(1, "," & strCompare & ",", "," & mrsSelItems!��Ŀ��� & ",") = 0 Then GoTo toPreCol
                    End If
                End If
                mrsPatient.Filter = 0
                mrsSelItems.Filter = 0
            End If
        Else
toPreRow:
            If vsf.ROW > 1 Then
                vsf.ROW = vsf.ROW - 1
                vsf.Col = vsf.Cols - 1
                GoTo toPreCol
            Else
                vsf.ROW = 1
            End If
            If vsf.RowHidden(vsf.ROW) Then GoTo toPreRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.ROW) = False Then
            vsf.TopRow = vsf.ROW
        End If
        Exit Sub
    End If
    
    mblnShow = True
    Call Vsf_EnterCell
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim introw As Integer, intCol As Integer
    If Button <> 1 Then Exit Sub
    
    intCol = vsf.MouseCol
    introw = vsf.MouseRow
    If introw = 0 And intCol = 0 Then
        Call vsf.Select(0, 0, vsf.Rows - 1, vsf.Cols - 1)
    ElseIf intCol = 0 Then
        Call vsf.Select(introw, 0, introw, vsf.Cols - 1)
    ElseIf introw = 0 Then
        Call vsf.Select(0, intCol, vsf.Rows - 1, intCol)
    End If
End Sub

Public Sub SignMe()
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim lngRecord As Long
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim introw As Integer, intRows As Integer
    On Error GoTo errHand
    '������ʱ��ѭ������ǩ��
    
    If mblnInit = False Then Exit Sub
    
    intRows = vsf.Rows - 1
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For introw = 1 To intRows
        If vsf.TextMatrix(introw, mlngSigner) = "" And vsf.TextMatrix(introw, vsf.Cols - 2) = gstrUserName Then
            If lngRecord <> Val(vsf.TextMatrix(introw, mlngRecord)) Then
                lngRecord = Val(vsf.TextMatrix(introw, mlngRecord))
                If SignName(introw, strSignTime, str״̬) = False Then Exit For
                blnSign = True
            End If
        End If
    Next
    
    If blnSign Then Call ShowMe(mfrmParent, mlng����ID, mstrPrivs, False, False)
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
    Dim introw As Integer, intRows As Integer
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mblnInit = False Then Exit Sub
    strSignTime = vsf.TextMatrix(vsf.ROW, mlngSignTime)
    blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    
    '��ͬһǩ��ʱ���������ȡ����,����ȡ��ǩ��
    gstrSQL = " Select A.����ID,A.��ҳID,A.Ӥ��,A.����ʱ��,B.��¼�� AS ǩ���� From ���˻����¼ A,���˻������� B" & _
              " Where A.ID=B.��¼ID And A.������Դ=2 And B.��¼����=5 And B.��Ŀ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strSignTime)
  
    With rsTemp
        Do While Not .EOF
            If rsTemp!ǩ���� = gstrUserName Then
                If UnSignName(!����ID, !��ҳID, !Ӥ��, Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), blnClear) = False Then Exit Sub
                blnUnSign = True
            End If
            .MoveNext
        Loop
    End With
    If blnUnSign Then Call ShowMe(mfrmParent, mlng����ID, mstrPrivs, False, False)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal lngRow As Long, ByVal strSignTime As String, str״̬ As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strStart As String
    Dim strSource As String
    Dim lngLoop As Long
    Dim lng����ID As Long, lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer
    
    On Error GoTo errHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    
    mrsPatient.Filter = "��=" & lngRow
    If mrsPatient.RecordCount = 0 Then Exit Function
    '�ȶ�λ�޸Ĺ�����,��������ѭ���ҵ��޸Ĺ�����
    lng����ID = mrsPatient!����ID
    lng��ҳID = mrsPatient!��ҳID
    lng����ID = mrsPatient!����ID
    intӤ�� = mrsPatient!Ӥ��
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5 And Nvl(a.��ʼ�汾,1)=Nvl(b.���汾,1)"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID, CDate(strStart), intӤ��)
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
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID, CDate(strStart), intӤ��)
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        MsgBox "��ǰû����Ҫǩ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    '76223:������,2014-08-05,����ǩ�����ʱ�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, lng����ID, lng��ҳID, mlng����ID, str״̬)
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���ӻ����¼_SignName("
        gstrSQL = gstrSQL & lng����ID & "," & lng��ҳID & "," & intӤ�� & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "',"
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "','" & oSign.ʱ�����Ϣ & "')"

        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckSigned(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, ByVal str����ʱ�� As String, _
    Optional ByRef lng֤��ID As Long, Optional ByRef strǩ���� As String, Optional ByRef strǩ��ʱ�� As String, Optional ByVal blnCheck As Boolean = True) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.��ĿID AS ֤��ID,A.��¼�� AS ǩ����,A.��Ŀ���� AS ǩ��ʱ�� From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID, CDate(str����ʱ��), intӤ��)
    If rs.BOF Then
        If blnCheck Then MsgBox "��ǰû����Ҫȡ����ǩ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng֤��ID = NVL(rs!֤��ID, 0)
    strǩ���� = NVL(rs!ǩ����)
    strǩ��ʱ�� = NVL(rs!ǩ��ʱ��)
    CheckSigned = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UnSignName(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, ByVal str����ʱ�� As String, ByVal blnClear As Boolean) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim lng֤��ID As Long
    Dim strSource As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    If Not CheckSigned(lng����ID, lng��ҳID, intӤ��, str����ʱ��, lng֤��ID) Then Exit Function
    
    '����ǵ���ǩ��,����Ҫ��֤
    '------------------------------------------------------------------------------------------------------------------
    If lng֤��ID > 0 Then
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
    gstrSQL = gstrSQL & lng����ID & ","
    gstrSQL = gstrSQL & lng��ҳID & ","
    gstrSQL = gstrSQL & intӤ�� & ","
    gstrSQL = gstrSQL & "To_Date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss')," & _
                      IIf(blnClear, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    UnSignName = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    mblnShow = False
    picInput.Visible = False
    SaveME = True
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lng����ID As Long, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnShow As Boolean = True)
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '���أ� ��
    '******************************************************************************************************************
'    Dim bln������ As Boolean
    
    Err = 0
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mstrPrivs = strPrivs
    mlng����ID = lng����ID
    Set mfrmParent = frmParent
    If mlng����ID = 0 Then Exit Sub

    mstrSel = ""
    mblnShow = False
    picInput.Visible = False

    Call ReadData
    
    mblnChange = False
    RaiseEvent AfterRefresh

    Call vsf_AfterRowColChange(2, 2, 1, 1)
    Call dkpMain.RecalcLayout
    
    '����ĳЩ�в��ƶ�
    vsf.FrozenCols = ��Ч������ - 1
    vsf.SheetBorder = &HC0C0FF
    
    If blnShow Then Me.Show 1, frmParent
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
                
                For lngRow = 1 To lngRows
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        StrText = vsf.TextMatrix(lngRow, lngCol)
                        If Trim(StrText) <> "" Then
                            If Not CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, strֵ��) Then
                                vsf.ROW = lngRow
                                If vsf.RowIsVisible(vsf.ROW) Then vsf.TopRow = vsf.ROW
                                Exit Function
                            End If
                        End If
                    End If
                Next
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

Private Function CheckValid(ByRef StrText As String, ByVal lngOrder As Long, ByVal lngClass As Long, ByVal strCap As String, _
    ByVal lngLength As Long, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal strֵ�� As String) As Boolean
    Dim arrData
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
    
    If lngClass = 2 Then '����ǻ��Ŀ����ܴ��ڲ�λ,�Ѳ�λ�����,ֻ���¼��������Ƿ񳬹�����
        If InStr(1, StrText, ":") <> 0 Then
            StrText = Split(StrText, ":")(1)
        End If
    End If
    
'    If strֵ�� = "" Then  '��ͨ��Ŀ
'        If Not (lngOrder = 9 Or lngOrder = 10) Then '���������ų�����������Ч��Χ���
'            If LenB(StrConv(strText, vbFromUnicode)) > lngLength Then
'                MsgBox "��" & lngRow & "�е�" & strCap & "���������飡", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'    Else                    '�������������Լ�Ѫѹ
        'û�����ʵ�ʱ�򣬲�����¼������
        If lngOrder = 2 And mbln���� Then
            If InStr(1, StrText, "/") <> 0 Then
                MsgBox "�뽫��õ���������¼�뵥�������ʵ�Ԫ���У�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If lngOrder = 3 Then
            If InStr(1, StrText, "/") <> 0 Then
                MsgBox "��������¼�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If lngOrder = 4 Or lngOrder = 5 Then
            'Ѫѹֵ���뺬/
            If vsf.TextMatrix(0, lngCol) Like "Ѫѹ*" Then
                If InStr(1, StrText, "/") = 0 Then
                    MsgBox "Ѫѹ���ݵĸ�ʽ��������ѹ/����ѹ��", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(Split(StrText, "/")(0)) = "" Or Trim(Split(StrText, "/")(1)) = "" Then
                    MsgBox "Ѫѹ���ݴ�������ѹ/����ѹ��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If UBound(Split(StrText, "/")) > 1 Then
            MsgBox "��" & lngRow & "�е�" & strCap & "����¼��������飡", vbInformation, gstrSysName
            Exit Function
        End If
        
        arrData = Split(StrText, "/")
        intCount = UBound(arrData)
        For intDo = 0 To intCount
            StrText = arrData(intDo)
            If InStr(1, StrText, ":") <> 0 Then StrText = Split(StrText, ":")(1)
            '������Ŀ������Ƿ񳬳�
            If lngOrder > 3 Then
                If LenB(StrConv(StrText, vbFromUnicode)) > lngLength Then
                    MsgBox "��" & lngRow & "�е�" & strCap & "���������飡", vbInformation, gstrSysName
                    vsf.TopRow = lngRow
                    Exit Function
                End If
            End If
            If IsNumeric(StrText) Then    '��Ч��Χ�뵱ǰ¼��ֵ������ֵ�Ͳż��,���򵱳���δ��˵��
                If Not (lngOrder = 9 Or lngOrder = 10) Then '���������ų�����������Ч��Χ���
                    If strֵ�� <> "" Then
                        If IsNumeric(Split(strֵ��, ";")(0)) Then
                            If Not (Val(StrText) >= Split(strֵ��, ";")(0) And Val(StrText) <= Split(strֵ��, ";")(1)) Then
                                MsgBox "��" & lngRow & "�е�" & strCap & "������Ч��Χ��" & Split(strֵ��, ";")(0) & "-" & Split(strֵ��, ";")(1) & "�������飡", vbInformation, gstrSysName
                                Exit Function
                            End If
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
        Next
'    End If
    
    'ƴװ���봮
    StrText = IIf(strPart <> "", strPart & ":", "") & strValue1 & IIf(strValue2 <> "", "/" & strValue2, "")
    CheckValid = True
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, blnOper As Boolean         'ָ��ĳ��ʱ������Ƿ��������
    Dim lngOrder As Long, lng����ID As Long, lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer
    Dim strTmp As String
    Dim intAllow As Integer, intType As Integer, lngClass As Long
    Dim str���� As String, str��� As String, str��λ As String, strδ��˵�� As String 'str���:ֻ�������⽵�»�������׾
    Dim lngRecord As Long, lngGroup As Long
    Dim lngRow As Long, lngRows As Long, lngCol As Long, lngCols As Long
    Dim strDate As String, strStart As String, strEnd As String, strSQLtmp As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim intPos As Integer, intMax As Integer
    Dim strSQL() As String
    On Error GoTo errHand
    'ͬһ��ʱ����(ͬһ����¼ID),��������ֶ�������,Ҳ����ֻ����һ��������������Ĵ���
    
    ReDim Preserve strSQL(1 To 1)
    lngRows = vsf.Rows - 1
    lngCols = mlngSigner - 1         '�����ǩ����,ǩ��ʱ��,��¼ID,��Ų�����
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '׼����������
    lng����ID = 0
    lng��ҳID = 0
    lng����ID = 0
    intӤ�� = 0
    For lngRow = 1 To lngRows
        mrsPatient.Filter = "��=" & lngRow
        If mrsPatient.RecordCount <> 0 Then
            '�ȶ�λ�޸Ĺ�����,��������ѭ���ҵ��޸Ĺ�����
            If lng����ID <> mrsPatient!����ID Or lng��ҳID <> mrsPatient!��ҳID Or intӤ�� <> mrsPatient!Ӥ�� Then
                lng����ID = mrsPatient!����ID
                lng��ҳID = mrsPatient!��ҳID
                lng����ID = mrsPatient!����ID
                intӤ�� = mrsPatient!Ӥ��
                blnOper = False
            End If
            
            strStart = Format(Me.dtp.Value, "yyyy-MM-dd HH:mm:ss")
            strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm:ss")
            
            '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
            If Val(vsf.RowData(lngRow)) = 1 Then
                If Not CheckTime(lngRow, lng����ID, lng��ҳID, Mid(strStart, 1, 16), Mid(strDate, 1, 16)) Then Exit Function
            End If
            
            '�����ȡ���ݺ��޸���ʱ��,��Ҫ����ʱ��
            If lngRecord <> Val(vsf.TextMatrix(lngRow, mlngRecord)) And Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 Then
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                gstrSQL = "Zl_���˻����¼_UpdateReplace(" & lngRecord & ",0," & intӤ�� & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
            
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
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID, intӤ��, lng����ID, CDate(strStart))
                    lngGroup = NVL(rsTemp!���, 0) + 1
                End If
                
                'һ��Ԫ��һ��Ԫ�صĴ���
                For lngCol = ��Ч������ To lngCols
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        
                        '�����ݽ������������޸Ĳ���
                        gstrSQL = "Zl_���˻����¼_UpdateRecord("
                        gstrSQL = gstrSQL & mrsPatient!����ID & "," & mrsPatient!��ҳID & "," & mrsPatient!Ӥ�� & ","
                        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & IIf(lngCol <> mlngOper, 1, 4) & ","
                        
                        lngOrder = 0
                        If lngCol <> mlngOper Then
                            mrsSelItems.Filter = "��=" & lngCol
                            mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
                            lngClass = mrsItems!��Ŀ����
                            intType = mrsItems!��Ŀ����
                            lngOrder = mrsItems!��Ŀ���
                        End If
                        strSQLtmp = gstrSQL
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
                            gstrSQL = gstrSQL & lng����ID & "," & lng��ҳID & "," & intӤ�� & ","
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
        If strSQL(intPos) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intPos), Me.Caption)
    Next
    SaveData = True
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnChange = False
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
    
    RaiseEvent AfterDataChanged
    Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
    lng����ID = Me.cbo����.ItemData(Me.cbo����.ListIndex)  '��������
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
    
    If lngRow = 0 Then lngRow = vsf.ROW
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ��¼+��ŵ����汾", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)))
    lng��߰汾 = NVL(rsTemp!��߰汾, 0)
    
    'ȡ��ǰ��Ŀ�ĵ�ǰ�汾
    gstrSQL = " Select MAX(��ʼ�汾) AS ��ǰ�汾 From ���˻������� Where ��¼ID=[1] And ��¼���=[2]" & IIf(lngCol = mlngOper, " And ��¼����=4", " And ��Ŀ���=[3]")
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
    ByVal strTime As String, ByVal strCurTime As String) As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    gstrSQL = " Select ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char( nvl(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
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
                MsgBox "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]", vbInformation, gstrSysName
                Exit Function
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
        .Filter = "��ʼԭ��=1"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 1 And strTime < !��ʼʱ�� Then
                MsgBox "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And strTime < !��ʼʱ�� Then
                MsgBox "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And strTime > !��ֹʱ�� Then
                MsgBox "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '�������˵��
        MsgBox "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]", vbInformation, gstrSysName
        GoTo exitHand
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
End Function
