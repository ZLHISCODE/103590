VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileOpen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����¼����"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTendFileOpen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   8205
      ScaleHeight     =   2610
      ScaleWidth      =   2730
      TabIndex        =   8
      Top             =   1245
      Visible         =   0   'False
      Width           =   2760
      Begin VSFlex8Ctl.VSFlexGrid vfgThisPrint 
         Height          =   1695
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Width           =   2340
         _cx             =   4128
         _cy             =   2990
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
         BackColorFixed  =   15790320
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
         Rows            =   4
         Cols            =   6
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      End
      Begin VB.TextBox txtLength 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1005
         Left            =   465
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label lblSubHeadPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblTitlePrint 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   75
         Width           =   1275
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   2670
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   979
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   979
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendFileOpen.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
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
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   4170
      Left            =   135
      TabIndex        =   1
      Top             =   1320
      Width           =   7500
      _cx             =   13229
      _cy             =   7355
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
      BackColorFixed  =   15790320
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
      Rows            =   4
      Cols            =   6
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   End
   Begin RichTextLib.RichTextBox rtbHead 
      Height          =   1200
      Left            =   2730
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   2117
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmTendFileOpen.frx":0E1C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbFoot 
      Height          =   1200
      Left            =   2730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6540
      Visible         =   0   'False
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   2117
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmTendFileOpen.frx":0EB9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThisRowHeight 
      Height          =   1695
      Left            =   8430
      TabIndex        =   13
      Top             =   4335
      Width           =   2340
      _cx             =   4128
      _cy             =   2990
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
      BackColorFixed  =   15790320
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
      Rows            =   4
      Cols            =   6
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   End
   Begin VB.Label lblSubhead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����:##"
      Height          =   180
      Left            =   315
      TabIndex        =   3
      Top             =   1050
      Width           =   630
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "һ�㻤���¼��"
      Height          =   180
      Left            =   3105
      TabIndex        =   2
      Top             =   510
      Width           =   1275
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   135
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTendFileOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'�������
'######################################################################################################################
Private mblnHead As Boolean, mblnFoot As Boolean
Private mlngPatiId As Long, mlngPageId As Long, mlngDeptId As Long, mintBaby As Integer
Private mstrPeriod As String
Private mbyt������ As Byte
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTitleFont As New StdFont, lngTitleFontSize As Long '����������ɫ
Private mobjSubFont  As New StdFont, lngSubFontSize As Long  '��������С
Private mobjTagFont As New StdFont, lngTagFontSize As Long   '������ʽ����
Private mobjTagFontPrint As New StdFont '��ӡʱ��������
Private mlngTagColor As Long        '������ʽ��ɫ
Private mdblRowHeightMin As Double     '�����С�߶�
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrSQL As String           '��֯���������ݲ�ѯ���
Private mlngFileID As Long
Private mbln����ʱ��ϲ� As Boolean
Private mblnʱ�������� As Boolean
Private mintStartCOLCount As Long, mintEndColCount As Long '��ʼ�̶������ͽ����̶���������ʼ�̶���Ϊ���̶���+���ڡ�ʱ�䣬�����̶���Ϊ��ʿ��ǩ���ˡ�ǩ��ʱ�䡢ǩ������
'��ʱ����
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar
Private mrsSumCol As New ADODB.Recordset
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private mblnStartUp As Boolean
Private strTemp As String
Private lngCurColor As Long, strCurFont As String, objFont As StdFont

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
Private Const gconLineHigh = 30

Public WithEvents zlEvent_Print As zlPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1
Public Event zlAfterPrint(ByVal lngFileID As Long)

Private mbytFontSize As Byte        '�����С0-9������,1-12������
'######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim lngCount As Long
    Dim aryItem() As String
    Dim blnTag As Boolean
    Dim lngReDraw As Long
    
    mobjSubFont.Size = lngSubFontSize
    Set CtlFont = mobjSubFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = BlowUp(CtlFont.Size)
    Set Me.Font = CtlFont
    '��������
    mobjTitleFont.Size = lngTitleFontSize
    Set CtlFont = mobjTitleFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = BlowUp(CtlFont.Size)
    lblTitle.AutoSize = True
    Set lblTitle.Font = CtlFont
    lblTitle.AutoSize = False
    '�ı�����
    Set lblSubhead.Font = Me.Font
    '�������
    With vfgThis
         Set .Font = Me.Font
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        .RowHeightMin = BlowUp(mdblRowHeightMin)
        '�п�����
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - 2), "`")(0)))
        Next lngCount
        
         '������ʽ
         mobjTagFont.Size = lngTagFontSize
        Set CtlFont = mobjTagFont
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = BlowUp(CtlFont.Size)
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If IsDate(.TextMatrix(lngCount, 0)) Then
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = CtlFont
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        '89729:������,���������С���������������ԣ��Զ��������߶�
        .AutoSizeMode = flexAutoSizeRowHeight
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        .Redraw = lngReDraw
    End With
    
    If mblnChildForm = False Then
        Set CtlFont = cbsThis.Options.Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = mbytFontSize
        Set cbsThis.Options.Font = CtlFont
        
        cbsThis.RecalcLayout
    Else
        Call cbsThis_Resize
    End If
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Private Function SumTend(ByVal lngFile As Long, ByVal lngPatientKey As Long, ByVal lngPageKey As Long) As Boolean
    '*****************************************************************************************************************
    '���ܣ� �������з�Χ�ڵĻ�������
    '������
    '���أ�
    '*****************************************************************************************************************
    Dim rsGroup As New ADODB.Recordset
    Dim strSQL As String
    Dim rsTimePeriod As New ADODB.Recordset
    
    On Error GoTo errHand

    mrsSumCol.Filter = ""
    If mrsSumCol.RecordCount = 0 Then Exit Function
    
    '���Ҳ��˵�ҽ����¼,�������⻤���������ܵĳ���ҽ��
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select a.��ʼִ��ʱ��,Nvl(a.ִ����ֹʱ��,Sysdate+365) As ִ����ֹʱ��,b.�걾��λ From ����ҽ����¼ a,������ĿĿ¼ b Where a.����id=[1] And a.��ҳid=[2] And b.ID=a.������Ŀid And a.ҽ��״̬ Not In (1,2,4) And b.��������='12' And b.���='Z' Order By a.��ʼִ��ʱ��"
    Set rsTimePeriod = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientKey, lngPageKey)
    If rsTimePeriod.BOF Then Exit Function
    
    'ͳ�Ƶ�ʱ���,���û��ʱ���,�˳������л���
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı� " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '����ʱ��' And d.�����ı� Is Not Null" & _
        " Order By d.�������, d.�����д�"
    Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFile)
    If rsGroup.BOF Then Exit Function
        
    With vfgThis
        
        Do While Not rsTimePeriod.EOF
                        
            Call SumRangeTend(rsGroup, Format(rsTimePeriod("��ʼִ��ʱ��").Value, "yyyy-MM-dd HH:mm:ss"), _
                                        Format(rsTimePeriod("ִ����ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss"), _
                                        Format(.TextMatrix(2, 1), "yyyy-MM-dd HH:mm:ss"), _
                                        Format(.TextMatrix(.Rows - 1, 1), "yyyy-MM-dd HH:mm:ss"), _
                                        zlCommFun.NVL(rsTimePeriod("�걾��λ").Value))
                        
            rsTimePeriod.MoveNext
        Loop

    End With
        
    SumTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function SumRangeTend(ByVal rsGroup As ADODB.Recordset, ByVal strStartTime As String, ByVal strEndTime As String, ByVal strMinTime As String, ByVal strMaxTime As String, Optional ByVal strColumn As String) As Boolean
    '*****************************************************************************************************************
    '���ܣ� ����ָ����Χ�ڵĻ�������
    '������
    '���أ�
    '*****************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim lngLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strSum As String
    Dim strTime As String
    Dim strSvrTime As String
    Dim rsResult As New ADODB.Recordset
    Dim lngLen As Long
    Dim lngRow As Long
    Dim intStartCol As Integer
    Dim intEndCol As Integer
    Dim strSumCol As String
    Dim lngDyas As Long
    Dim str��ֹʱ�� As String
    Dim str��ʼʱ�� As String
    Dim lngStartRow As Long, lngEndRow As Long
    Dim blnAllow As Boolean
    
    On Error GoTo errHand
    
    
    If strMaxTime < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") Then strMaxTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    
    '��ʼ������
    '------------------------------------------------------------------------------------------------------------------
    Set rsResult = New ADODB.Recordset
    With rsResult
        .Fields.Append "�к�", adBigInt
        .Fields.Append "ʱ��", adVarChar, 30
        .Fields.Append "���", adVarChar, 100
        .Fields.Append "����", adVarChar, 100
        .Open
        
    End With
    
    mrsSumCol.Filter = ""
    mrsSumCol.MoveFirst
    
    '��ʱ�䷶Χȷ�Ͽ�ʼ��,������
    '------------------------------------------------------------------------------------------------------------------
    With vfgThis
        For lngLoop = 3 To .Rows - 1    '��ͷ�̶���3��,�п��������˲��̶ֹ���,������:mintTabTiers
            strTime = Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss")
            If lngStartRow = 0 Then
                If strTime >= strStartTime Then
                    lngStartRow = lngLoop
                End If
            End If
            
            If lngEndRow = 0 Then
                If strTime > strEndTime Then
                    lngEndRow = lngLoop - 1
                    Exit For
                End If
            End If
        Next
        If lngLoop = .Rows Then
            If lngEndRow = 0 Then lngEndRow = .Rows - 1
        End If
    End With
    If Not (lngStartRow > 0 And lngStartRow <= lngEndRow) Then Exit Function
    
    '��ͳ�Ƶ�
    '------------------------------------------------------------------------------------------------------------------
    
    lngDyas = DateDiff("d", CDate(strStartTime), CDate(strEndTime))
    rsGroup.MoveFirst
    Do While Not rsGroup.EOF
        strTmp = zlCommFun.NVL(rsGroup!�����ı�)
        If strTmp <> "" Then
            aryTmp = Split(strTmp, ",")
            If UBound(aryTmp) >= 2 Then
                
                If InStr(aryTmp(1), ":") = 0 Then aryTmp(1) = aryTmp(1) & ":00"
                If InStr(aryTmp(2), ":") = 0 Then aryTmp(2) = aryTmp(2) & ":00"
                
                For lngLoop = 0 To lngDyas
                    
                    str��ʼʱ�� = Format(DateAdd("d", lngLoop, CDate(strStartTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                    
                    If str��ʼʱ�� < strStartTime Then
                        strSvrTime = strStartTime
                    Else
                        strSvrTime = str��ʼʱ��
                    End If
                    
                    If Format(aryTmp(1), "HH:mm:ss") < Format(aryTmp(2), "HH:mm:59") Then
                        'ͬһ��
                        blnAllow = True
                        str��ֹʱ�� = Format(DateAdd("d", lngLoop, CDate(strStartTime)), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                    Else
                        '����ͬһ��
                        blnAllow = False
                        str��ֹʱ�� = Format(strSvrTime, "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                        strSvrTime = Format(DateAdd("d", -1, CDate(strSvrTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                    End If
                                                            
                    If str��ֹʱ�� > strMaxTime Then Exit For
                                                            
                    If str��ֹʱ�� >= strStartTime And str��ֹʱ�� <= strEndTime And str��ֹʱ�� >= strMinTime Then
                        
                        mrsSumCol.Filter = ""
                        If strColumn <> "" Then mrsSumCol.Filter = "����='" & strColumn & "'"
                        If mrsSumCol.RecordCount > 0 Then
                            mrsSumCol.MoveFirst
                            Do While Not mrsSumCol.EOF
                                
                                rsResult.AddNew
                                rsResult("�к�").Value = mrsSumCol("�к�").Value
                                rsResult("ʱ��").Value = str��ֹʱ��
                                rsResult("���").Value = 0
                                
                                lngLen = DateDiff("n", CDate(strSvrTime), CDate(str��ֹʱ��))
                                strSum = Format(lngLen \ 60, "00") & "Сʱ" & Format(lngLen Mod 60, "00") & "��"
                                
                                If blnAllow Then
                                    rsResult("����").Value = Format(str��ֹʱ��, "MM-dd") & " " & aryTmp(0) & "(" & strSum & ")"
                                Else
                                    rsResult("����").Value = Format(strSvrTime, "MM-dd") & " " & aryTmp(0) & "(" & strSum & ")"
                                End If
                                
                                mrsSumCol.MoveNext
                            Loop
                        End If
                    End If
                Next
            End If
        End If
        rsGroup.MoveNext
    Loop
    
    '������л��ܣ�����д����Ӧ��rsResult��¼����
    '------------------------------------------------------------------------------------------------------------------
    strSum = ""
    rsGroup.MoveFirst
    With vfgThis
        .AddItem "": vfgThisPrint.AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(DateAdd("d", 1, CDate(.TextMatrix(.Rows - 2, 1))), "yyyy-MM-dd") & " 23:59:59"
        vfgThisPrint.TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
        lngEndRow = lngEndRow + 1
        
        '������ϸ���ɻ���������
        '--------------------------------------------------------------------------------------------------------------
        Do While Not rsGroup.EOF
            
            strTmp = zlCommFun.NVL(rsGroup!�����ı�)
            If strTmp <> "" Then
                aryTmp = Split(strTmp, ",")
                If UBound(aryTmp) >= 2 Then
                    
                    mrsSumCol.Filter = ""
                    If strColumn <> "" Then mrsSumCol.Filter = "����='" & strColumn & "'"
                    If mrsSumCol.RecordCount > 0 Then
                        mrsSumCol.MoveFirst
                        Do While Not mrsSumCol.EOF
    
                            If InStr(aryTmp(1), ":") = 0 Then aryTmp(1) = aryTmp(1) & ":00"
                            If InStr(aryTmp(2), ":") = 0 Then aryTmp(2) = aryTmp(2) & ":00"
                                        
                            For lngLoop = lngStartRow To lngEndRow
                                                                                        
                                strTime = Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss")
                                
                                If strTime > strEnd And strEnd <> "" Then
                                    '��д
                                    rsResult.Filter = ""
                                    rsResult.Filter = "ʱ��='" & strEnd & "' And �к�=" & mrsSumCol("�к�").Value & " And ���� Like '*" & Split(rsGroup!�����ı�, ",")(0) & "*'"
                                    If rsResult.RecordCount = 0 Then
                                        rsResult.AddNew
                                        rsResult("�к�").Value = mrsSumCol("�к�").Value
                                        rsResult("ʱ��").Value = strEnd
                                    End If
                                    rsResult("���").Value = Val(strSum)
                                        
                                    strSum = ""
                                    strStart = ""
                                    strEnd = ""
                                    strSvrTime = ""
                                End If
                                
                                'ȷ�������ʱ���
                                If Format(aryTmp(1), "HH:mm:ss") < Format(aryTmp(2), "HH:mm:59") Then
                                    'ʱ�䷶Χ��ͬһ������
                                    
                                    If (strStart = "" And strEnd = "") Or Not (strTime >= strStart And strTime <= strEnd) Then
                                        
                                        '�ж��ϴκ͵�ǰ�Ƿ���ͬһ��,�������,˵���м��������ݵ�ͳ��
                                        If strSvrTime <> "" And strTime <> "" Then
                                            If CDate(strSvrTime) <> CDate(strTime) Then
                                                strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                            End If
                                        End If
                                        
                                        strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                        strStart = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strEnd = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                                        If strTime > strEnd Then
                                            strStart = ""
                                            strEnd = ""
                                        End If
    
                                    End If
                                    
                                    If strTime >= strStart And strTime <= strEnd And strEnd <> "" Then
                                        strSum = Val(strSum) + Val(.TextMatrix(lngLoop, mrsSumCol("�к�").Value))
                                    End If
                                
                                Else
                                    'ʱ�䷶Χ����ͬһ������,������һ��
                                    
                                    If (strStart = "" And strEnd = "") Or Not (strTime >= strStart Or strTime <= strEnd) Then
                                        
'                                        strEnd = Format(CDate(strTime), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
'                                        strStart = Format(DateAdd("d", -1, CDate(strTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                        strStart = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strEnd = Format(DateAdd("d", 1, CDate(strTime)), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
    
                                    End If
                                                                    
                                    If (strTime >= strStart And strTime <= strEnd) And strEnd <> "" Then
                                        strSum = Val(strSum) + Val(.TextMatrix(lngLoop, mrsSumCol("�к�").Value))
                                    End If
                                
                                End If
                                    
                            Next
                        
                            If strEnd <> "" Then
                                rsResult.Filter = ""
                                rsResult.Filter = "ʱ��='" & strEnd & "' And �к�=" & mrsSumCol("�к�").Value & " And ���� Like '*" & Split(rsGroup!�����ı�, ",")(0) & "*'"
                                If rsResult.RecordCount = 0 Then
                                    rsResult.AddNew
                                    rsResult("�к�").Value = mrsSumCol("�к�").Value
                                    rsResult("ʱ��").Value = strEnd
                                End If
                                rsResult("���").Value = Val(strSum)
                                    
                                strSum = ""
                                strStart = ""
                                strEnd = ""
                                strSvrTime = ""
                            End If
                            
                            mrsSumCol.MoveNext
                        Loop
                    End If
                End If
            End If
            
            rsGroup.MoveNext
        Loop
        
        rsResult.Filter = "���>0"
        Do While Not rsResult.EOF
            Debug.Print rsResult!ʱ�� & ","; rsResult!��� & "," & rsResult!����
            rsResult.MoveNext
        Loop
        
        '���������ɵĻ������ݽ��в�����ʾ
        '--------------------------------------------------------------------------------------------------------------
        Call ShowSumTend(rsResult, lngStartRow, lngEndRow, strColumn)
        
    End With
        
    SumRangeTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Function

Private Function ShowSumTend(ByVal rsResult As ADODB.Recordset, ByVal lngStartRow As Long, ByVal lngEndRow As Long, Optional ByVal strColumn As String) As Boolean
    '*****************************************************************************************************************
    '���ܣ� ��ʾ��������
    '������
    '���أ�
    '*****************************************************************************************************************
    Dim lngRow As Long
    Dim lngLoop As Long
    Dim intEndCol As Integer
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim intStartCol As Integer
    
    On Error GoTo errHand
    If rsResult.RecordCount = 0 Then Exit Function
    
    rsResult.MoveFirst
    rsResult.Sort = "ʱ��"
    
    mrsSumCol.Filter = ""
    If strColumn <> "" Then mrsSumCol.Filter = "����='" & strColumn & "'"
    If mrsSumCol.RecordCount = 0 Then Exit Function
    
    mrsSumCol.Sort = "�к�"
    If mbln����ʱ��ϲ� = False Then
        intStartCol = 2
    Else
        intStartCol = 1
    End If
    intEndCol = mrsSumCol("�к�").Value
    If intEndCol <= 2 Then Exit Function

    With vfgThis
        
        lngLoop = lngStartRow
        
        Do While Not rsResult.EOF

            If Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss") > Format(rsResult("ʱ��").Value, "yyyy-MM-dd HH:mm:ss") And .Cell(flexcpData, lngLoop, 1, lngLoop, 1) = "" Then
                If .Cell(flexcpData, lngLoop, 1, lngLoop, 1) <> rsResult("����").Value Then
                    
                    If .Cell(flexcpData, lngLoop - 1, 1, lngLoop - 1, 1) <> rsResult("����").Value Then
                        .AddItem "", lngLoop
                        
                        .MergeRow(lngLoop) = True
                        .Cell(flexcpText, lngLoop, intStartCol, lngLoop, intEndCol) = rsResult("����").Value
                        .Cell(flexcpAlignment, lngLoop, intStartCol, lngLoop, intEndCol) = flexAlignCenterCenter
                        .Cell(flexcpData, lngLoop, 1, lngLoop, 1) = rsResult("����").Value
                        .Cell(flexcpForeColor, lngLoop, 0, lngLoop, .Cols - 1) = 255
'                        .Cell(flexcpForeColor, lngLoop, intStartCol, lngLoop, intEndCol) = 255
                        
                        vfgThisPrint.AddItem "", lngLoop
                        
                        vfgThisPrint.MergeRow(lngLoop) = True
                        vfgThisPrint.Cell(flexcpText, lngLoop, intStartCol, lngLoop, intEndCol) = rsResult("����").Value
                        vfgThisPrint.Cell(flexcpAlignment, lngLoop, intStartCol, lngLoop, intEndCol) = flexAlignCenterCenter
                        vfgThisPrint.Cell(flexcpData, lngLoop, 1, lngLoop, 1) = rsResult("����").Value
                        vfgThisPrint.Cell(flexcpForeColor, lngLoop, 0, lngLoop, .Cols - 1) = 255
                        
                        lngEndRow = lngEndRow + 1
                        
                        lngRow = lngLoop
                    Else
                        lngRow = lngLoop - 1
                    End If
                    
                    If rsResult("�к�").Value Mod 2 = 1 Then
                        .TextMatrix(lngRow, rsResult("�к�").Value) = rsResult("���").Value
                    Else
                        .TextMatrix(lngRow, rsResult("�к�").Value) = " " & rsResult("���").Value
                    End If
                    
                    .Cell(flexcpAlignment, lngRow, rsResult("�к�").Value, lngRow, rsResult("�к�").Value) = .Cell(flexcpAlignment, 2, rsResult("�к�").Value, 2, rsResult("�к�").Value)
                    vfgThisPrint.TextMatrix(lngRow, rsResult("�к�").Value) = .TextMatrix(lngRow, rsResult("�к�").Value)
                    vfgThisPrint.Cell(flexcpAlignment, lngRow, rsResult("�к�").Value, lngRow, rsResult("�к�").Value) = .Cell(flexcpAlignment, lngRow, rsResult("�к�").Value, lngRow, rsResult("�к�").Value)
                End If
                
                rsResult.MoveNext
            Else
                lngLoop = lngLoop + 1
                If lngLoop > lngEndRow Then Exit Do
            End If

        Loop

        
        .Rows = .Rows - 1
        
        '�������һ������(���һ����������ռ�ʱ�䷶Χ��ֹͣ��,��û��ȫ���ܽ�)
        '--------------------------------------------------------------------------------------------------------------
        strTmp = ""
        For lngLoop = .Rows - 1 To 1 Step -1
            If .Cell(flexcpData, lngLoop - 1, 1, lngLoop - 1, 1) = "" Then
                Exit For
            ElseIf .Cell(flexcpData, lngLoop, 1, lngLoop, 1) <> "" Then
                strTmp = strTmp & "," & lngLoop
            End If
        Next

        If strTmp <> "" Then
            aryTmp = Split(strTmp, ",")
            For lngLoop = 0 To UBound(aryTmp)
                If Val(aryTmp(lngLoop)) > 0 Then
                    .RemoveItem Val(aryTmp(lngLoop))
                    vfgThisPrint.RemoveItem Val(aryTmp(lngLoop))
                End If
            Next
        End If
    End With
    
    ShowSumTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    
End Function

Public Function zlPrintTend(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDeviceName As String) As Boolean
    
    '1-Ԥ��,2-��ӡ
    
    Select Case bytMode
    Case 1
        Call zlRptPrint(2, strPrintDeviceName)
    Case 2
        Call zlRptPrint(1, strPrintDeviceName)
    Case 3
        Call zlRptPrint(3, strPrintDeviceName)
    End Select
   
    '
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, lngDeptId As Long, Optional ByVal intBaby As Integer = 0, Optional ByVal strPeriod As String, Optional ByVal blnChildForm As Boolean = False, Optional ByVal byt������ As Byte = 3, Optional ByVal blnDataMoved As Boolean, Optional ByVal bytSize As Byte = 0)
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngFileID           �����ļ���ʽ���
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       bytSize             '�����С 0-9������ ��1-12������
    '���أ� ��
    '******************************************************************************************************************
'    Dim bln������ As Boolean
    
    Err = 0
    Dim stdObjFont As StdFont
    Dim rsItem As New ADODB.Recordset
    Dim rsCol As New ADODB.Recordset
    
    On Error GoTo errHand
    mlngFileID = lngFileID
    
    mblnChildForm = blnChildForm
    mblnMoved_HL = blnDataMoved
    If mblnChildForm = False Then
        Call InitForm
    Else
        
        If mblnStartUp Then
            'Me.WindowState = 2
            Call FormSetCaption(Me, False, False)
            
            stbThis.Visible = Not mblnChildForm
            cbsThis.ActiveMenuBar.Visible = False
            cbsThis.RecalcLayout
            mblnStartUp = False
        End If
        
    End If
    
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    picPrint.Visible = False
    lblSubHeadPrint.Caption = "": lblSubHeadPrint.Tag = ""
    
    Set mrsSumCol = New ADODB.Recordset
    With mrsSumCol
        .Fields.Append "�к�", adBigInt
        .Fields.Append "����", adVarChar, 50
        .Fields.Append "�б���", adVarChar, 100
        .Open
    End With
    
    '65164:������,2013-08-27
    Set rsCol = New ADODB.Recordset
    With rsCol
        .Fields.Append "���", adBigInt
        .Open
    End With
    '��ȡ���ܻ�����Ŀ
    gstrSQL = "Select ��Ŀ���� From �����¼��Ŀ where ��Ŀ����=0 And ��Ŀ��ʾ=4"
    Call zlDatabase.OpenRecordset(rsItem, gstrSQL, "��ȡ���ܻ�����Ŀ")
    
    '������ʽ��ȡ
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select l.���� From �����ļ��б� l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    If mblnChildForm = False Then Me.Caption = "�����¼���� - " & rsTemp!����
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mlngDeptId = lngDeptId
    mintBaby = intBaby
    mstrPeriod = strPeriod
    mbyt������ = byt������
    mbln����ʱ��ϲ� = False
    mblnʱ�������� = False
    mintStartCOLCount = 0
    mintEndColCount = 0
'    bln������ = (Val(zlDatabase.GetPara("�����������", glngSys, 1255, "0")) = 1)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":  Me.vfgThis.Cols = Val("" & !�����ı�): Me.vfgThisPrint.Cols = Me.vfgThis.Cols
            Case "��С�и�": Me.vfgThis.RowHeightMin = Val("" & !�����ı�): mdblRowHeightMin = Me.vfgThis.RowHeightMin: Me.vfgThisPrint.RowHeightMin = Me.vfgThis.RowHeightMin
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.vfgThis.Font = objFont
                Set Me.lblSubhead.Font = Me.vfgThis.Font
                Set Me.Font = Me.lblSubhead.Font
                Set mobjSubFont = objFont
                lngSubFontSize = objFont.Size
                
                Set stdObjFont = New StdFont
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.lblSubHeadPrint.Font = stdObjFont
                Set Me.vfgThisPrint.Font = stdObjFont
                Set Me.picPrint.Font = stdObjFont
            Case "�ı���ɫ": Me.vfgThis.ForeColor = Val("" & !�����ı�): Me.vfgThisPrint.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ"
                Me.vfgThis.GridColor = Val("" & !�����ı�): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                Me.vfgThisPrint.GridColor = Val("" & !�����ı�): Me.vfgThisPrint.GridColorFixed = Me.vfgThis.GridColor
            Case "�����ı�": Me.lblTitle.Caption = "" & !�����ı�: Me.lblTitlePrint.Caption = Me.lblTitle.Caption
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                Set stdObjFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.lblTitle.Font = objFont
                Me.lblTitle.AutoSize = False
                Set mobjTitleFont = objFont
                lngTitleFontSize = objFont.Size
                
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.lblTitlePrint.Font = stdObjFont
                Me.lblTitlePrint.AutoSize = False
                
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
                lngTagFontSize = objFont.Size
                
                Set stdObjFont = New StdFont
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFontPrint = stdObjFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "����ʱ��ϲ�": mbln����ʱ��ϲ� = (Val("" & !�����ı�) = 1)
            '65502:������,2013-11-12
            Case "ʱ��������": mblnʱ�������� = (Val("" & !�����ı�) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    If mblnʱ�������� = True Then mbln����ʱ��ϲ� = False
    
    
    gstrSQL = "Select ����||'-'||��� AS KEY,��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
        mblnHead = ReadPageHead(rtbHead, rsTemp!Key)
        mblnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        
        mbyt������ = rsTemp!����
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, strSql�� As String, strSql�� As String, strSql�� As String, strSQL���� As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim lngColumn As Long, lngNum As Long
    'lngNum ������Ϊ��Ŀ������ð��,��������㲻�ܶ�����,������滻Ϊ������
    lngNum = 1

    
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    '65164:������,2013-08-27
    With rsTemp
        Do While Not .EOF
            rsCol.AddNew
            rsCol("���") = Val(NVL(!�������, 0))
            rsCol.Update
        .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    
    With rsTemp
        lngColumn = 0: mstrColWidth = ""
        strSql�� = "": strSql�� = "": strSql�� = "": strSql�� = "": strSQL���� = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            
            If lngColumn <> !������� Then
                mstrColWidth = mstrColWidth & "," & !��������
                If NVL(!Ҫ������) <> "" Then
                    If strSql�� <> "" Then
                        strSql�� = strSql�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        strSql�� = strSql�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                Else
                    If strSql�� <> "" Then
                        strSql�� = strSql�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        strSql�� = strSql�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                lngColumn = !�������
            End If
            
            '53172:������,2013-04-25,�޸���ȡ��ʿ��c.��¼�˸�Ϊl.������
            Select Case NVL(!Ҫ������)
            Case "����"
                bln���� = True
                strSql�� = strSql�� & ",����"
                strSql�� = strSql�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                strSql�� = strSql�� & ",ǩ����"
                strSql�� = strSql�� & ",a.��¼�� As ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                strSql�� = strSql�� & ",ǩ��ʱ��"
                strSql�� = strSql�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ������"
                blnǩ������ = True
                strSql�� = strSql�� & ",ǩ������"
                strSql�� = strSql�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����, 1,11)) As ǩ������"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ʱ��"
                blnʱ�� = True
                strSql�� = strSql�� & ",ʱ��"
                strSql�� = strSql�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "��ʿ"
                bln��ʿ = True
                strSql�� = strSql�� & ",��ʿ"
                'strSql�� = strSql�� & ",c.��¼�� As ��ʿ"
                strSql�� = strSql�� & ",l.������ as ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If NVL(!Ҫ������) <> "" Then
                    strSql�� = strSql�� & ",Max(""" & !Ҫ������ & """) As """ & "B" & Format(lngNum, "00") & """"
                    
                    strSQL���� = strSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    strSql�� = strSql�� & "||""B" & Format(lngNum, "00") & """"
                    
                    If Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "" Then
                        strSql�� = strSql�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
'                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                    Else
                        strSql�� = strSql�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,Null,'" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
'                        strSql�� = strSql�� & "||Decode(""" & !Ҫ������ & """, Null, Null,'" & !�����ı� & "'||""" & !Ҫ������ & """||'" & !Ҫ�ص�λ & "')"
                    End If
                    lngNum = lngNum + 1
                End If
            End Select
            
            '65164:������,2013-08-27,28�汾ǰ��Ϊû�л�����Ŀ��ʶ������ʽ��������Ҫִ�л��������ж��Ƿ������Ŀ(ֻ���ÿ�аﶥһ����Ŀ)
            '28�汾�����¼��Ŀ�����˻��ܱ�ʶ��ĿǰֻҪ�ǻ����ж����л��ܡ�
            rsItem.Filter = "��Ŀ����='" & "" & !Ҫ������ & "'"
            rsCol.Filter = "���=" & lngColumn
            'If zlCommFun.NVL(!Ҫ�ر�ʾ, 0) = 1 Then
            If rsItem.RecordCount > 0 And rsCol.RecordCount = 1 Then
                mrsSumCol.Filter = ""
                mrsSumCol.Filter = "�к�=" & lngColumn + 1
                If mrsSumCol.RecordCount = 0 Then
                    mrsSumCol.AddNew
                    mrsSumCol("�к�") = lngColumn + 1
                    mrsSumCol("����") = "" & !Ҫ������
                    mrsSumCol("�б���") = "" & !Ҫ������
                    mrsSumCol.Update
                End If
            End If
            
            .MoveNext
        Loop
        If Mid(strSql��, 3) <> "" Then
            strSql�� = strSql�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            strSql�� = strSql�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If strSQL���� <> "" Then strSQL���� = "(" & Mid(strSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then strSql�� = strSql�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then strSql�� = strSql�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then strSql�� = strSql�� & ",l.������ as ��ʿ"
        'If bln��ʿ = False Then strSql�� = strSql�� & ",c.��¼�� As ��ʿ"
        
        If blnǩ���� = False Then strSql�� = strSql�� & ",a.��¼�� As ǩ����"
        If blnǩ������ = False Then strSql�� = strSql�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,1,11)) As ǩ������"
        If blnǩ��ʱ�� = False Then strSql�� = strSql�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
        
        If Mid(strSql��, 2) = "" Then
            ShowSimpleMsg "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡"
            Exit Sub
        End If
        If bln���� = True Then mintStartCOLCount = mintStartCOLCount + 1
        If blnʱ�� = True Then mintStartCOLCount = mintStartCOLCount + 1
        If bln��ʿ = True Then mintEndColCount = mintEndColCount + 1
        
        If blnǩ���� = True Then mintEndColCount = mintEndColCount + 1
        If blnǩ������ = True Then mintEndColCount = mintEndColCount + 1
        If blnǩ��ʱ�� = True Then mintEndColCount = mintEndColCount + 1
        
        mintStartCOLCount = mintStartCOLCount + 2
        mstrSQL = "Select ����,����ʱ��," & Mid(strSql��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(strSql��, 2) & vbCrLf & _
                "        From (Select c.��¼���,����ʱ��," & Mid(strSql��, 2) & vbCrLf & _
                "               From ���˻����¼ l, ���˻������� c,���˻������� a " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.����id = [1] And l.��ҳid = [2] And a.��¼id(+)=l.ID And a.��¼����(+)=5 And a.��ֹ�汾(+) IS NULL And Nvl(l.Ӥ��,0)=[4] And c.��ֹ�汾 Is Null And c.��¼����<>5 And l.����id + 0 = [3] And l.����ʱ�� Between [5] And [6] And l.������<=[7])" & vbCrLf & _
                IIf(strSQL���� <> "", "Where " & strSQL����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��" & _
                                "       Order By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��)"
                
        mstrColWidth = Mid(mstrColWidth, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Call zlRefresh
    '������ʾ
    If blnChildForm = False Then
        Call SetFontSize(bytSize)
        If frmParent Is Nothing Then
            Me.Show vbModal
        Else
            Me.Show vbModal, frmParent
        End If
        
        Unload Me
    Else
        '102173:���Ӳ�������ӡ�����¼��(���˴������ݻ�����),��ӡ�ڶ�����������
        mblnStartUp = False
'        Call cbsThis_Resize
    End If
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitForm() As Boolean
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����Ԫ����̬����
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Function

Private Sub zlRefresh(Optional ByVal blnReSize As Boolean = False)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, strCell As String
    Dim blnTag As Boolean
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strTmp As String
    '�п�����
    Dim blnAlign As Boolean
    Dim dblWidth As Double  '������ʱ���еĿ��
    Dim strCol As String
    
    Err = 0: On Error GoTo errHand
    
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    lblSubHeadPrint.Caption = ""
    lblSubHeadPrint.Tag = ""
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubhead, "|")
    
    aryPeriod = Split(mstrPeriod, "��")
    aryPeriod(0) = Format(aryPeriod(0) & ":00", "yyyy-MM-dd HH:mm:ss")
    aryPeriod(1) = Format(aryPeriod(1) & ":59", "yyyy-MM-dd HH:mm:ss")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        Select Case strItemName
        Case "��ǰ����"
        
            strTmpSQL = "Select b.����" & vbNewLine & _
                        "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                        "Order By a.��ʼʱ��"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            
        Case "��ǰ����"
        
            strTmpSQL = "Select a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "��ǰ����"
        
            strTmpSQL = "Select ���� From ���ű� a Where a.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngDeptId)
            
        Case "סԺҽʦ"
            strTmpSQL = "Select a.����ҽʦ" & vbNewLine & _
                        "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "���λ�ʿ"
        
            strTmpSQL = "Select a.���λ�ʿ" & vbNewLine & _
                        "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "����ȼ�"
'            ��֪����,ע����
'            strTmpSQL = "Select b.����" & vbNewLine & _
'                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
'                        "            From ���˱䶯��¼" & vbNewLine & _
'                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,������ĿĿ¼ b" & vbNewLine & _
'                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.����ȼ�ID Is Not Null And b.ID=a.����ȼ�ID" & vbNewLine & _
'                        "Order By a.��ʼʱ��"
                        
            strTmpSQL = "Select b.����" & vbNewLine & _
                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [5] Between a.��ʼʱ�� And a.��ֹʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case Else
            strTmp = ""
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strPrefix, strItemName, mlngPatiId, mlngPageId, mintBaby)
        End Select
        
        If rsTemp.BOF = False Then
            If strTmp <> "" Then
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & rsTemp.Fields(0).Value
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    lblSubHeadPrint.Tag = lblSubhead.Tag
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    
    'װ������
    gstrSQL = mstrSQL
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, mintBaby, CDate(aryPeriod(0)), CDate(aryPeriod(1)), mbyt������)
    
    '����ţ�51746�������ɣ�2012-06-18 15:16�����������С
    '��ӡ���
    With Me.vfgThisPrint
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '��ͷ��д
        '65164:������,2013-08-27,�޸ĺϲ���ʽ
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        strCol = ""
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
            
            If mbln����ʱ��ϲ� And InStr(1, ",����,ʱ��,", "," & strCell & ",") > 0 And strCell <> "" Then
                .ColHidden(lngCol + 1) = True
                strCol = strCol & "," & lngCol + 1
            End If
            If strCell = "ʱ��" And mblnʱ�������� = True Then .ColHidden(lngCol + 1) = True
        Next
        
        '�п�����
        blnAlign = False
        dblWidth = 0 '������ʱ���еĿ��
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
            If mbln����ʱ��ϲ� And InStr(1, strCol & ",", "," & lngCount & ",") > 0 Then
                dblWidth = dblWidth + .ColWidth(lngCount)
            End If
            If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                blnAlign = True
                .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
            End If
        Next
        '������ʱ������ʾ����,�п�Ϊ������ʱ���е��ܿ��
        If mbln����ʱ��ϲ� Then
            .ColHidden(1) = False
            .ColWidth(1) = dblWidth
            .TextMatrix(0, 1) = "����ʱ��"
            If mintTabTiers >= 2 Then .TextMatrix(1, 1) = "����ʱ��"
            If mintTabTiers >= 3 Then .TextMatrix(2, 1) = "����ʱ��"
        End If
        
        '������ʽ
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If mintTagFormHour < mintTagToHour Then
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            Else
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFontPrint
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        
        '����Ӧ�߶Ⱥ�������ʽ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeRowHeight
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = 0
'            .ROWHEIGHT(2) = 0
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
        
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = 0
            
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
            
            
        Case 3
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = .RowHeightMin

            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        .Redraw = flexRDDirect
    End With
    
    With Me.vfgThis
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '��ͷ��д
        '65164:������,2013-08-27,�޸ĺϲ���ʽ
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        strCol = ""
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
            
            If mbln����ʱ��ϲ� And InStr(1, "����,ʱ��", strCell) > 0 And strCell <> "" Then
                .ColHidden(lngCol + 1) = True
                strCol = strCol & "," & lngCol + 1
            End If
            If strCell = "ʱ��" And mblnʱ�������� = True Then .ColHidden(lngCol + 1) = True
        Next
        
        '�п�����
        blnAlign = False
        dblWidth = 0 '������ʱ���еĿ��
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
            If mbln����ʱ��ϲ� And InStr(1, strCol & ",", "," & lngCount & ",") > 0 Then
                dblWidth = dblWidth + .ColWidth(lngCount)
            End If
            If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                blnAlign = True
                vfgThis.ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
            End If
        Next
        '������ʱ������ʾ����,�п�Ϊ������ʱ���е��ܿ��
        If mbln����ʱ��ϲ� Then
            .ColHidden(1) = False
            .ColWidth(1) = dblWidth
            .TextMatrix(0, 1) = "����ʱ��"
            If mintTabTiers >= 2 Then .TextMatrix(1, 1) = "����ʱ��"
            If mintTabTiers >= 3 Then .TextMatrix(2, 1) = "����ʱ��"
        End If
        
        '������ʽ
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If mintTagFormHour < mintTagToHour Then
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            Else
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        
        '����Ӧ�߶Ⱥ�������ʽ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeRowHeight
        '�ٰ��кϲ�
        For lngCount = 0 To vfgThis.Cols - 1
            vfgThis.MergeCol(lngCount) = True
        Next
        vfgThis.AutoSize 0, vfgThis.Cols - 1
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = 0
'            .ROWHEIGHT(2) = 0
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
        
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = 0
            
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
            
            
        Case 3
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = .RowHeightMin

            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        Call SumTend(mlngFileID, mlngPatiId, mlngPageId)
        
        .Redraw = flexRDDirect
    End With
    
    '��ģ̬������ˢ����Ҫ�������������С
    If blnReSize = True Then Call ReSetFontSize
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub aa()

End Sub
Private Sub zlLableBruit()
    
    '���ݿ�ȷ�ɢ��ǩ
    Dim aryRow() As String
    Dim lngSpaces As Long
'    aryRow = Split(Me.lblSubhead.Tag, vbCrLf)
'
'    For lngCount = 0 To UBound(aryRow)
'        If UBound(Split(aryRow(lngCount), Space(1))) > 0 Then
'            lngSpaces = 1
'            Do
'                If Me.TextWidth(Join(Split(aryRow(lngCount), Space(1)), Space(lngSpaces + 1))) > Me.vfgThis.Width Then
'                    If lngSpaces > 1 Then lngSpaces = lngSpaces - 1
'                    aryRow(lngCount) = Join(Split(aryRow(lngCount), Space(1)), Space(lngSpaces))
'                    Exit Do
'                End If
'                lngSpaces = lngSpaces + 1
'            Loop
'        End If
'    Next
'    Me.lblSubhead.Caption = Join(aryRow, vbCrLf)
    Me.lblSubhead.Caption = Me.lblSubhead.Tag
    Me.lblSubHeadPrint.Caption = Me.lblSubHeadPrint.Tag
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Me.vfgThis.Move lngScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    Me.vfgThis.Height = lngScaleBottom - Me.vfgThis.Top - 210
End Sub

Private Sub VsfToMsh(Optional ByVal blnShow As Boolean = False)
    Dim dblHeight As Double
    Dim lngRow As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    On Error Resume Next
    
    '1����ת����ͷ
    lngRows = vfgThisPrint.FixedRows - 1
    lngCols = vfgThisPrint.Cols - 1
    '���ñ�ͷ������ʽ
    mshHead.Rows = lngRows + 2
    mshHead.FixedRows = lngRows + 1
    mshHead.Cols = lngCols + 1
    mshHead.FixedCols = vfgThisPrint.FixedCols
    mshHead.MergeCells = flexMergeFree
    mshHead.ROWHEIGHT(mshHead.Rows - 1) = 0
    Set mshHead.Font = vfgThisPrint.Font
    For lngRow = 0 To lngRows
        mshHead.Row = lngRow
        vfgThisPrint.Row = lngRow
        If vfgThisPrint.RowHidden(lngRow) Then
            mshHead.ROWHEIGHT(lngRow) = 0
        Else
            mshHead.ROWHEIGHT(lngRow) = vfgThisPrint.ROWHEIGHT(lngRow)
        End If
        dblHeight = dblHeight + vfgThisPrint.ROWHEIGHT(lngRow)
        For lngCol = 0 To lngCols
            mshHead.Col = lngCol
            vfgThisPrint.Col = lngCol
            mshHead.CellAlignment = vfgThisPrint.CellAlignment
            mshHead.TextMatrix(lngRow, lngCol) = vfgThisPrint.TextMatrix(lngRow, lngCol)
            If lngRow = lngRows Then '�����п�
                If vfgThisPrint.ColHidden(lngCol) Then
                    mshHead.ColWidth(lngCol) = 0
                Else
                    mshHead.ColWidth(lngCol) = vfgThisPrint.ColWidth(lngCol)
                End If
            End If
        Next
        mshHead.MergeRow(lngRow) = True
    Next
    '�����кϲ�
    For lngCol = 0 To lngCols
        mshHead.MergeCol(lngCol) = True
    Next
    
    '2����ת������
    lngRows = vfgThisPrint.Rows - vfgThisPrint.FixedRows
    mshDetail.Rows = lngRows
    mshDetail.Cols = lngCols + 1
    mshDetail.FixedRows = 0
    mshHead.FixedCols = vfgThisPrint.FixedCols
    mshDetail.WordWrap = False
    '65164:������,2013-08-27,�޸Ĵ�ӡ���ܺϲ�����
    mshDetail.MergeCells = flexMergeFree
    Set mshDetail.Font = vfgThisPrint.Font
    Set picPrint.Font = vfgThisPrint.Font
    Set vfgThisRowHeight.Font = vfgThisPrint.Font
    
    For lngRow = 0 To lngRows - 1
        mshDetail.Row = lngRow
        If vfgThisPrint.RowHidden(lngRow + vfgThisPrint.FixedRows) Then
            mshDetail.ROWHEIGHT(lngRow) = 0
        Else
            mshDetail.ROWHEIGHT(lngRow) = vfgThisPrint.ROWHEIGHT(lngRow + vfgThisPrint.FixedRows)
        End If
        For lngCol = 0 To lngCols
            mshDetail.Col = lngCol
            vfgThisPrint.Row = lngRow + vfgThisPrint.FixedRows
            vfgThisPrint.Col = lngCol
            mshDetail.CellForeColor = vfgThisPrint.CellForeColor
            
            mshDetail.CellAlignment = vfgThisPrint.ColAlignment(lngCol)
            
            mshDetail.TextMatrix(lngRow, lngCol) = vfgThisPrint.TextMatrix(lngRow + vfgThisPrint.FixedRows, lngCol)
            If lngRow = lngRows - 1 Then '�����п�
                If vfgThisPrint.ColHidden(lngCol) Then
                    mshDetail.ColWidth(lngCol) = 0
                Else
                    mshDetail.ColWidth(lngCol) = vfgThisPrint.ColWidth(lngCol)
                End If
            End If
        Next
        '65164:������,2013-08-27,�޸Ĵ�ӡ���ܺϲ�����
        mshDetail.MergeRow(lngRow) = vfgThisPrint.MergeRow(lngRow + vfgThisPrint.FixedRows)
    Next
    
    '���ô�С��λ��
    mshHead.Move vfgThisPrint.Left, vfgThisPrint.Top, vfgThisPrint.Width, dblHeight
    mshDetail.Move vfgThisPrint.Left, vfgThisPrint.Top + dblHeight, vfgThisPrint.Width, vfgThisPrint.Height - dblHeight
    
    mshHead.Visible = blnShow
    mshDetail.Visible = blnShow
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal strPrintDeviceName As String)
    Dim objPrint As New zlPrint2Grd, objAppRow As zlTabAppRow
    Dim lngWidth As Long, lngEmptyLR As Long, lngScaleWidth As Long
    On Error GoTo errHand
    
    If zlEvent_Print Is Nothing Then
        Set zlEvent_Print = VBA.GetObject("", "zl9PrintMode.zlPrintMethod")
    End If
    
    objPrint.EmptyUp = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageUp", 20)
    objPrint.EmptyDown = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageDown", 20)
    objPrint.EmptyLeft = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageLeft", 20)
    objPrint.EmptyRight = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageRight", 20)
        
    '���ô�ӡ��ʽ
    If mblnHead Then mstrPageHead = Me.rtbHead.Text
    If mblnFoot Then mstrPageFoot = Me.rtbFoot.Text
    SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageHead", mstrPageHead
    SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageFoot", mstrPageFoot
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    lngEmptyLR = 0
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then
        lngEmptyLR = lngEmptyLR + Val(Split(mstrPaperSet, ";")(4))
        objPrint.EmptyLeft = Round(Me.ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageLeft", objPrint.EmptyLeft
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then
        lngEmptyLR = lngEmptyLR + Val(Split(mstrPaperSet, ";")(5))
        objPrint.EmptyRight = Round(Me.ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageRight", objPrint.EmptyRight
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then
        objPrint.EmptyUp = Round(Me.ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageUp", objPrint.EmptyUp
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then
        objPrint.EmptyDown = Round(Me.ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "PageDown", objPrint.EmptyDown
    End If
    
    '84140:LPF,���ϱ�ǩ�����������ݴ�ӡ���ü���
    On Error Resume Next
    Printer.PaperSize = Val(Split(mstrPaperSet, ";")(0))
    Printer.Orientation = Val(Split(mstrPaperSet, ";")(1))
    
    If Printer.PaperSize = 256 Then
        Call SetCustonPager(Me.hWnd, Val(Split(mstrPaperSet, ";")(3)), Val(Split(mstrPaperSet, ";")(2)))
    End If
    lngScaleWidth = Printer.Width - lngEmptyLR
    On Error GoTo errHand

    Call VsfToMsh(False)
    
    Set objPrint.BodyHead = Me.mshHead
    Set objPrint.BodyGrid = Me.mshDetail
    objPrint.Title.Text = lblTitlePrint.Caption
    Set objPrint.Title.Font = lblTitlePrint.Font
    Set objPrint.AppFont = lblSubHeadPrint.Font
    
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngStart As Long, lngPos As Long, lngMAX As Long, lngNumber As Long, blnNumber As Boolean, lngAsc As Long
    lngSpaces = lblSubHeadPrint.Height / 210
    strLable = lblSubHeadPrint.Caption
    lngMAX = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMAX
        '�����ѧ����,��������Ƶ���һ����ʾ
        lngAsc = Asc(Mid(strLable, lngPos, 1))

        '����Ƿ񳬿�(���ȳ����п�,���������س����з�)
        If picPrint.TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "��") > lngScaleWidth Or lngPos = lngMAX Or lngAsc = 10 Then

            strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            lngStart = lngPos + 1
            
            '���������
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    Next
    
    lngWidth = Val(Split(mstrPaperSet, ";")(3))
    If mstrPageHead <> "" Then objPrint.Header = mstrPageHead
    If mstrPageFoot <> "" Then
        mstrPageFoot = Replace(mstrPageFoot, "{��ӡʱ��}", Now)
        mstrPageFoot = Replace(mstrPageFoot, "{��ӡ��}", gstrUserName)
        objPrint.Footer = mstrPageFoot ' LeftB(mstrPageFoot & Space(lngWidth), lngWidth - objPrint.EmptyLeft - objPrint.EmptyRight)
    End If
    
    
    If bytMode = 1 Then
        If strPrintDeviceName = "" Then
            bytMode = zlEvent_Print.zlPrintAsk(objPrint)
        Else
            SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", strPrintDeviceName
        End If
        
        objPrint.Footer = mstrPageFoot
        Call ReSetTableRows(objPrint)
        If bytMode <> 0 Then zlEvent_Print.zlPrintOrView2Grd objPrint, bytMode
    Else
        Call ReSetTableRows(objPrint)
        zlEvent_Print.zlPrintOrView2Grd objPrint, bytMode
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strItemKey As String
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call zlRefresh(True)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With Me.lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 120
    End With
    Me.vfgThis.Move lngScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    Me.vfgThis.Height = lngScaleBottom - Me.vfgThis.Top - 210
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfgThis.Rows > Me.vfgThis.FixedRows)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If mblnChildForm = False Then Call SaveWinState(Me, App.ProductName)
    Set mobjTagFont = Nothing
    Set cbrControl = Nothing
    Set cbrMenuBar = Nothing
    Set cbrToolBar = Nothing
    Set mrsSumCol = Nothing
    Set rsTemp = Nothing
    Set objFont = Nothing
    Set zlEvent_Print = Nothing
        
    Set mobjTitleFont = Nothing
    Set mobjSubFont = Nothing
    Set mobjTagFontPrint = Nothing
    mblnStartUp = False
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vfgThis.AutoSize 0, vfgThis.Cols - 1
End Sub

Private Sub zlEvent_Print_zlAfterPrint()
    RaiseEvent zlAfterPrint(mlngFileID)
End Sub

Private Function ReadPageHead(objHead As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

'-------------------------------------------------------------------
'66724: ������,2014-1-15
'���ܣ����ݿؼ��ĳ��ȼ����ı�ռ�õ�����
Private Function GetData(ByVal strInput As String, Optional ByVal strSplit As String = "'") As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", strSplit) & strData
    Next
    GetData = Split(GetData, strSplit)
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '��Ϊ��,��ʾ�������ַ���������
    Next
    strLine(1) = 1
End Sub

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function


Private Function ReSetTableRows(objsend As zlPrint2Grd) As Boolean
'----------------------------------------------------------------------------------------
'���ܣ�����ֽ�Ŵ�С�����������������������������
'objSend:��ӡ����
'����ʱ����Ԥ����ӡ֮ǰ����(zlRptPrint)
'˵�����ú���������㷨���մ�ӡ�����Ĵ�����Ч������������������������ܱ������㼰����
'----------------------------------------------------------------------------------------
    Dim sgnTitle As Single, sgnUpAppRow As Single, sgnDownAppRow As Single, sgnFixRow As Single '����߶ȣ�������Ŀ�߶�,������Ŀ�߶ȣ����̶��и߶�
    Dim sgnHeight As Single '���岿����Ч����߶�
    Dim lngRow As Long, lngCol As Long, lngStartRow As Long
    Dim sgnTmpHeight As Single, sgnRHeight As Single, sgnTextHeight As Single
    Dim arrData, intDatas As Integer, intData As Integer, i As Integer
    Dim arrColText() As String, sgnRowHeight As Single, StrText As String, sgnRowHeightNew As Single
    Dim lngNum As Long '�������������ڼ�¼����������
    Dim sgnRowHeightCurrent As Single  '��ǰ���ʵ�ʸ߶�  �����102102
    On Error GoTo errHand
    
    'һ:����ʵ������������Ч�߶�
    If Not zlGetPrinterSet Then Exit Function
    Set picPrint.Font = objsend.Title.Font
    sgnTitle = picPrint.TextHeight(objsend.Title.Text) + 2 * gconLineHigh
    Set picPrint.Font = objsend.AppFont
    sgnUpAppRow = (picPrint.TextHeight("jg") + gconLineHigh) * objsend.UnderAppRows.Count + gconLineHigh
    sgnDownAppRow = (picPrint.TextHeight("jg") + gconLineHigh) * objsend.BelowAppRows.Count + gconLineHigh
    
    For lngRow = 0 To Me.mshHead.FixedRows - 1
        sgnFixRow = sgnFixRow + Me.mshHead.ROWHEIGHT(lngRow)
    Next lngRow
    sgnHeight = Printer.ScaleHeight - (objsend.EmptyUp + objsend.EmptyDown) * conRatemmToTwip - sgnTitle - sgnUpAppRow - sgnDownAppRow - sgnFixRow - 2 * gconLineHigh
    
    If sgnHeight < vfgThisPrint.RowHeightMin Then ReSetTableRows = True: Exit Function
    
    '����ѭ��������ݼ���Ƿ񳬳���Χ������������ڱ�����׷���µ�һ�д�ų�������
    sgnTmpHeight = 0
    lngStartRow = 0
    lngNum = 0
    Set picPrint.Font = mshDetail.Font
PreForStart:
    For lngRow = lngStartRow To Me.mshDetail.Rows - 1
PreBegin:
        '�����в�����
        If mshDetail.MergeRow(lngRow) = False Then
            ReDim arrColText(0 To mshDetail.Cols - 1)
            If sgnTmpHeight + Me.mshDetail.ROWHEIGHT(lngRow) > sgnHeight Then
                sgnRHeight = sgnHeight - sgnTmpHeight '������ĸ߶�
                If sgnRHeight >= vfgThisPrint.RowHeightMin Then
                    '��ʱ����֮��ʼ
                    sgnRowHeight = 0
                    For lngCol = mintStartCOLCount To mshDetail.Cols - mintEndColCount - 1
                            '��ȡ�ı�����ռ�õ�����
                            arrColText(lngCol) = ""
                            With txtLength
                                .Width = mshDetail.ColWidth(lngCol)
                                .Text = Replace(Replace(Replace(mshDetail.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                .FontName = mshDetail.CellFontName
                                .FontSize = mshDetail.CellFontSize
                                .FontBold = mshDetail.CellFontBold
                                .FontItalic = mshDetail.CellFontItalic
                            End With
                            arrData = GetData(txtLength.Text)
                            intDatas = UBound(arrData)
                            '���㲢��¼�������ı�����
                            sgnRowHeightCurrent = 0
                            sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intDatas, arrData)
                            If intDatas > 0 And sgnRowHeightCurrent > sgnRHeight Then
                                sgnTextHeight = 0
                                For intData = 0 To intDatas
                                    sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intData, arrData)
                                    If sgnRowHeightCurrent > sgnRHeight Then
                                        If intData = 0 Then GoTo PreEnd
                                        StrText = ""
                                        For i = 0 To intData - 1
                                            StrText = StrText & arrData(i)
                                        Next i
                                        mshDetail.TextMatrix(lngRow, lngCol) = StrText
                                        sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intData - 1, arrData)
                                        If sgnRowHeight < sgnRowHeightCurrent Then
                                            sgnRowHeight = sgnRowHeightCurrent
                                        End If
                                        For i = intData To intDatas
                                            arrColText(lngCol) = arrColText(lngCol) & arrData(i)
                                        Next i
                                        Exit For
                                    End If
                                Next intData
                            End If
                    Next lngCol
                    'sgnRowHeight > 0˵�����ݳ����˿�����ķ�Χ
                    If sgnRowHeight > 0 Then
                        If sgnRowHeight < vfgThisPrint.RowHeightMin Then sgnRowHeight = vfgThisPrint.RowHeightMin
                        mshDetail.ROWHEIGHT(lngRow) = sgnRowHeight
                    End If
                    sgnRowHeight = 0
                    '���㳬�����ݵ����߶�
                    For lngCol = mintStartCOLCount To mshDetail.Cols - mintEndColCount - 1
                        If arrColText(lngCol) <> "" Then
                            With txtLength
                                .Width = mshDetail.ColWidth(lngCol)
                                .Text = Replace(Replace(Replace(arrColText(lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                .FontName = mshDetail.CellFontName
                                .FontSize = mshDetail.CellFontSize
                                .FontBold = mshDetail.CellFontBold
                                .FontItalic = mshDetail.CellFontItalic
                            End With
                            
                            arrData = GetData(txtLength.Text)
                            intDatas = UBound(arrData)
                            sgnRowHeightNew = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intDatas, arrData)
                            If sgnRowHeight < sgnRowHeightNew Then
                                sgnRowHeight = sgnRowHeightNew
                            End If
                            
                        End If
                    Next lngCol
                    '��ɱ���е���Ӻ͸�ֵ
                    If sgnRowHeight > 0 Then
                        '���ݿ�ҳ�󣬱�����������ҳ��Ҫ��ʾ���ڡ�ʱ�䡢��ʿ��ǩ���ˡ�ǩ��ʱ�䡢ǩ������
                        For i = 0 To mintStartCOLCount - 1
                            arrColText(i) = mshDetail.TextMatrix(lngRow, i)
                        Next i
                        '��ʿ��ǩ���ˡ�ǩ�����ڡ�ǩ��ʱ��ĸ�ֵ
                        For i = 0 To mintEndColCount - 1
                            arrColText(mshDetail.Cols - i - 1) = mshDetail.TextMatrix(lngRow, mshDetail.Cols - i - 1)
                        Next i
                        
                        If sgnRowHeight < vfgThisPrint.RowHeightMin Then sgnRowHeight = vfgThisPrint.RowHeightMin
                        mshDetail.AddItem "", lngRow + 1: lngNum = lngNum + 1
                        mshDetail.ROWHEIGHT(lngRow + 1) = sgnRowHeight
                        On Error Resume Next
                        For lngCol = 0 To mshDetail.Cols - 1
                            mshDetail.Row = lngRow + 1: mshDetail.Col = lngCol
                            vfgThisPrint.Row = lngRow + vfgThisPrint.FixedRows - (lngNum - 1)
                            vfgThisPrint.Col = lngCol
                            mshDetail.CellForeColor = vfgThisPrint.CellForeColor
                            mshDetail.CellAlignment = vfgThisPrint.ColAlignment(lngCol)
                            mshDetail.TextMatrix(lngRow + 1, lngCol) = arrColText(lngCol)
                        Next lngCol
                        mshDetail.MergeRow(lngRow + 1) = mshDetail.MergeRow(lngRow)
                        If Err <> 0 Then Err.Clear
                        On Error GoTo errHand
                    End If
                    sgnTmpHeight = 0
                    lngStartRow = lngRow + 1
                    GoTo PreForStart
                Else
PreEnd:
                    sgnTmpHeight = 0
                    GoTo PreBegin
                End If
            Else
                sgnTmpHeight = sgnTmpHeight + Me.mshDetail.ROWHEIGHT(lngRow)
            End If
        End If
    Next lngRow
    
    ReSetTableRows = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zlGetCurrentVSFHight(ByVal ColWith As Long, ByVal Count As Long, ByVal arrData As Variant) As Single
'----------------------------------------------------------------------------------------
'���ܣ������������ݻ�ȡ��ǰ���и�
'colwith �п�  count ��ǰ�������� arrData ��������
'----------------------------------------------------------------------------------------
'102102
    Dim StrText As String
    Dim intData  As Long, intDatas As Long
    Dim i As Long
    vfgThisRowHeight.ColWidth(0) = ColWith
    vfgThisRowHeight.WordWrap = True
    vfgThisRowHeight.AutoSizeMode = flexAutoSizeRowHeight
    If Count < 0 Then Exit Function
    StrText = ""
    For i = 0 To Count
        StrText = StrText & arrData(i)
    Next i
    vfgThisRowHeight.TextMatrix(0, 0) = StrText
    vfgThisRowHeight.AutoSize 0, 0
    zlGetCurrentVSFHight = vfgThisRowHeight.ROWHEIGHT(0)
End Function

Private Function zlGetPrinterSet() As Boolean
    '------------------------------------------------
    '���ܣ���ȡ��ϵͳע���Ĵ�ӡȱʡ����
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        zlGetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
    If Printer.DeviceName <> strDeviceName Then
        For iCount = 0 To Printers.Count - 1
            If Printers(iCount).DeviceName = strDeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        Next
    End If
    
    Err = 0
    On Error Resume Next
    Printer.PaperBin = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Orientation", Printer.Orientation)
    
    intPaperSize = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Width", Printer.Width)
        lngHeight = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Height", Printer.Height)
        
        Call SetCustonPager(Me.hWnd, lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If

    zlGetPrinterSet = True
End Function

'-----------------------------------------------------------------------------------------------------
