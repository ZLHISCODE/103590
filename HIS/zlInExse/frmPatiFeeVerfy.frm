VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiFeeVerfy 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   270
      ScaleHeight     =   2955
      ScaleWidth      =   6240
      TabIndex        =   11
      Top             =   5355
      Width           =   6240
      Begin VB.PictureBox picImgList 
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   90
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   30
         Width           =   210
         Begin VB.Image imgColList 
            Height          =   195
            Index           =   0
            Left            =   0
            Picture         =   "frmPatiFeeVerfy.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Bindings        =   "frmPatiFeeVerfy.frx":054E
         Height          =   1395
         Left            =   15
         TabIndex        =   12
         Top             =   0
         Width           =   7110
         _cx             =   12541
         _cy             =   2461
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
         BackColorSel    =   16444122
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFeeVerfy.frx":0562
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picҽ�� 
      BorderStyle     =   0  'None
      Height          =   3765
      Left            =   330
      ScaleHeight     =   3765
      ScaleWidth      =   7245
      TabIndex        =   9
      Top             =   1170
      Width           =   7245
      Begin VB.PictureBox picImgList 
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   120
         Width           =   210
         Begin VB.Image imgColList 
            Height          =   195
            Index           =   1
            Left            =   0
            Picture         =   "frmPatiFeeVerfy.frx":059E
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3555
         Left            =   45
         TabIndex        =   10
         Top             =   90
         Width           =   5925
         _cx             =   10451
         _cy             =   6271
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
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFeeVerfy.frx":0AEC
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin MSComctlLib.ImageList img16 
            Left            =   1920
            Top             =   600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":0B87
                  Key             =   "ǩ��"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":0ED9
                  Key             =   "���δ�ӡ"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":1473
                  Key             =   ""
                  Object.Tag             =   "3"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img16dbl 
            Left            =   2535
            Top             =   615
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":1A0D
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1230
      ScaleHeight     =   600
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   240
      Width           =   11880
      Begin VB.CheckBox chk���� 
         Caption         =   "ֻ�����ʷ��õ�ҽ��"
         Height          =   180
         Left            =   7695
         TabIndex        =   1
         Top             =   165
         Width           =   1935
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&T)"
         Height          =   300
         Index           =   1
         Left            =   6615
         TabIndex        =   2
         Top             =   120
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&L)"
         Height          =   300
         Index           =   0
         Left            =   5640
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboCons 
         Height          =   300
         Index           =   1
         Left            =   2895
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   165
         Width           =   1875
      End
      Begin VB.ComboBox cboCons 
         Height          =   300
         Index           =   0
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "ִ�п���"
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "��Ч"
         Height          =   180
         Index           =   2
         Left            =   5115
         TabIndex        =   6
         Top             =   195
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imgFlag 
      Left            =   0
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":20A7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":22C1
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":27DB
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":29F5
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":2F0F
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3429
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3643
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3BDD
            Key             =   "���ͨ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4177
            Key             =   "���δͨ��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPass 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4711
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4FFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":52F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatiFeeVerfy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long, mlng��ҳID As Long
Private mlngModule As Long
'------------------------------------------------------------------
'�ֲ�����
Private Enum mEM_Pancel
    EM_���� = 1
    EM_ҽ�� = 2
    EM_���� = 3
End Enum
Private mrsSkinTest As ADODB.Recordset
Private mrsDefine As ADODB.Recordset    'ҽ�����ݶ���
Private mblnUnload As Boolean
Private mblnDataMove As Boolean '�Ƿ���ʷ��������
Private mlngFontSize As Long '�ֺŴ�С
Private mstr������� As String
Private mstrִ�п���ID As String
Private mrs������� As ADODB.Recordset
Private Enum CboIdx
    EM_IDX������� = 0
    EM_IDXִ�п��� = 1
End Enum
Private mblnNotClick As Boolean
Private mrsҽ�� As ADODB.Recordset
Private mblnChangeData As Boolean
Private mstrPrivs As String
Private mbytFontSize As Byte
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("��") + 20
        Case UCase("VsFlexGrid")
            Call zlControl.VSFSetFontSize(objCtrl, mbytFontSize)
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = CtlFont.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    Call picFilter_Resize
End Sub

Public Function ShowData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnDataMove As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ����
    '����:���˺�
    '����:2012-05-31 11:01:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    mblnDataMove = blnDataMove
    Call Loadҽ��
    ShowData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2012-05-30 13:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_����, 200, 580, DockTopOf, Nothing)
    panThis.Title = "��������": panThis.Handle = picFilter.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.MaxTrackSize.Height = picFilter.Height \ Screen.TwipsPerPixelY
    panThis.MinTrackSize.Height = picFilter.Height \ Screen.TwipsPerPixelY
    panThis.Tag = mEM_Pancel.EM_����
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_ҽ��, 250, 580, DockBottomOf, panThis)
    panThis.Title = "": panThis.Handle = picҽ��.hWnd
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = mEM_Pancel.EM_ҽ��
    
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_����, 250, 580, DockBottomOf, panThis)
    panThis.Title = "": panThis.Handle = picFeeList.hWnd
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = mEM_Pancel.EM_����
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

Private Sub cboCons_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
     Call Loadҽ��(True)
End Sub

Private Sub chkType_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        mblnNotClick = True
        chkType(Index).Value = 1
        mblnNotClick = False
    End If
    Loadҽ�� True
End Sub

Private Sub chk����_Click()
    If mblnNotClick Then Exit Sub
    Loadҽ�� True
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mEM_Pancel.EM_����
        Item.Handle = picFilter.hWnd
    Case mEM_Pancel.EM_ҽ��
        Item.Handle = picҽ��.hWnd
    Case mEM_Pancel.EM_����
        Item.Handle = picFeeList.hWnd
    End Select
End Sub
 
Private Function Loadҽ��(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ������
    '����:���˺�
    '����:2012-05-30 14:20:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    Dim strSQL As String, strWhere As String, lngҽ��ID As Long, lng���ID As Long
    Dim bytҽ����Ч As Byte, str���� As String, dt���� As Date
    Dim str���� As String, strFeeTable As String
    Dim strFilter As String
    
    Call InitAdvice
    If mlng����ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
    strFilter = ""
    On Error GoTo ErrHand:
    With vsAdvice
        If .Row > 0 Then
            i = .ColIndex("ҽ��ID")
            If i > 0 Then
                lngҽ��ID = Val(.TextMatrix(.Row, i))  '��¼��ǰ��
            End If
        End If
    End With
    If Not (chkType(0).Value = 1 And chkType(1).Value = 1) Then
        strFilter = strFilter & " And ��Ч='" & IIf(chkType(0).Value = 1, "����", "����") & "'"
    End If
    
    'ֻ��ʾ����δ���ʷ��õ�ҽ��
    If chk����.Value = 1 Then
         strFilter = strFilter & " And ����ҽ��ID<>0"
      '  str���� = _
            " And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From (Select Nvl(C.���id, C.ID) As ҽ��id" & vbNewLine & _
            "              From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C" & vbNewLine & _
            "              Where A.ҽ��id = C.ID And A.NO = B.NO And A.��¼���� = B.��¼���� And A.��¼���� = 2 And B.��¼״̬ = 0 And" & vbNewLine & _
            "                    C.����id = [1] And C.��ҳid = [2]" & ")" & vbNewLine & _
            "       Where A.ID = ҽ��id Or A.���id = ҽ��id )"
    End If
    With cboCons(CboIdx.EM_IDX�������)
        If .ListIndex >= 0 Then
            If Chr(.ItemData(.ListIndex)) <> 0 Then
                strFilter = strFilter & " and �������='" & Chr(.ItemData(.ListIndex)) & "'"
            End If
        End If
    End With
    
    With cboCons(CboIdx.EM_IDXִ�п���)
        If .ListIndex >= 0 Then
            If .ItemData(.ListIndex) <> 0 Then
                strFilter = strFilter & " And ִ�п���ID=" & .ItemData(.ListIndex) & ""
            End If
        End If
    End With
    
    strFeeTable = "" & _
        "   Select nvl(B.���ID,B.ID) as ҽ�����,Sum(nvl(Ӧ�ս��,0)) as Ӧ�ս��,Sum(nvl(ʵ�ս��,0)) as ʵ�ս��  " & _
        "   From סԺ���ü�¼ A,����ҽ����¼ B " & _
        "   Where A.ҽ�����=B.ID and  A.����ID=[1] and A.��ҳID=[2] " & _
        "   Group by nvl(B.���ID,B.ID)"
    strFeeTable = strFeeTable & " Union All " & Replace(strFeeTable, "סԺ���ü�¼", "������ü�¼")
    
    'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨'�������÷�����
    strSQL = _
    "   Select /*+ RULE */ A.ID as ҽ��ID,A.���ID,A.���,Nvl(A.Ӥ��,0) as Ӥ��ID,A.ҽ��״̬," & _
    "               Nvl(A.�������,'*') as �������,B.��������,C.�������,A.������־ as ��־,nvl(�Ƿ�������,0) as ���," & _
    "               A.�����,Decode(Nvl(A.ҽ����Ч,0),0,'����','����') as ��Ч," & _
    "               To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,A.ҽ������,Null as ����,A.Ƥ�Խ�� as Ƥ��," & _
    "               Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ)," & _
    "               '4',A.�ܸ�����||G.���㵥λ,'5',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,'6',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,A.�ܸ�����||B.���㵥λ)) as ����," & _
    "               Decode(A.��������,NULL,NULL,A.��������||Decode(A.�������,'4',G.���㵥λ,B.���㵥λ)) as ����,A.����," & _
    "               A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('2468',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�," & _
    "               A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��," & _
    "               nvl(E.ID,decode(nvl(A.ִ������,0),0,-1,5,-2,NULL)) as ִ�п���ID," & _
    "               Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
    "               Decode(Instr('567E',Nvl(A.�������,'*')),0,NULL,A.ִ������) as ִ������," & _
    "               To_Char(A.�ϴ�ִ��ʱ��,'YYYY-MM-DD HH24:MI') as �ϴ�ִ��," & _
    "               Decode(A.ҽ��״̬,1,'�¿�',2,'����',3,'У��',4,'����',5,'����',6,'��ͣ',7,'����',8,'ֹͣ',9,'ȷ��ֹͣ') as ״̬," & _
    "               A.����ҽ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,A.У�Ի�ʿ,To_Char(A.У��ʱ��,'YYYY-MM-DD HH24:MI') as У��ʱ��," & _
    "               A.ͣ��ҽ��,To_Char(A.ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ͣ��ʱ��,F.������Ա as ͣ����ʿ," & _
    "               To_Char(A.ȷ��ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ȷ��ͣ��ʱ��,A.������ĿID,B.�Թܱ���,A.ִ�б��,A.���δ�ӡ,A.ǰ��ID,Decode(S.ǩ��ID,NULL,0,1) as ǩ����," & _
    "               M.�����ļ�ID as �ļ�ID,Nvl(N.ͨ��,0) as ������,Y.����ID as ����ID,Y.����״̬,A.�շ�ϸĿID,B.���㵥λ as ������λ,A.��������ID,A.���״̬, " & _
    "               A.�������,A1.Ӧ�ս��,A1.ʵ�ս��, nvl(A1.ҽ�����,0) as ����ҽ��ID"
    strSQL = strSQL & _
    " From ����ҽ����¼ A,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,�շ���ĿĿ¼ G," & _
    "       ����ҽ��״̬ F,����ҽ��״̬ S,����ҽ������ Y,��������Ӧ�� M,�����ļ��б� N," & _
    "      (" & strFeeTable & ") A1" & _
    " Where A.������ĿID=B.ID(+) And nvl(a.���ID,a.Id) =A1.ҽ�����(+) And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+)" & _
    "       And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=G.ID(+) And A.ID=Y.ҽ��ID(+)" & _
    "       And (Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL) Or A.�������='E' And B.��������='8')" & _
    "       And A.ID=F.ҽ��ID(+) And F.��������(+)=9 And A.ID=S.ҽ��ID And S.��������=1" & _
    "       And A.������ĿID=M.������ĿID(+) And M.Ӧ�ó���(+)=2 And M.�����ļ�ID=N.ID(+) And N.����(+)=7" & _
    "       And A.����ID=[1] And A.��ҳID=[2] And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1"
    
    '������ʾ��ʽ����
    strSQL = strSQL & " Order by Ӥ��ID,���"
    
    '������ʷ�ռ䴦��
    If mblnDataMove Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
    End If
    If Not blnFilter Or mrsҽ�� Is Nothing Or mblnChangeData Then
        Set mrsҽ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID, bytҽ����Ч, dt����)
        With mrsҽ��
            mstr������� = "": mstrִ�п���ID = ""
            Do While Not .EOF
                If InStr(1, mstr������� & ",", "," & Nvl(!�������) & ",") = 0 And Nvl(!�������) <> "" Then
                    mstr������� = mstr������� & "," & Nvl(!�������)
                End If
                If InStr(1, mstrִ�п���ID & ",", "," & Val(Nvl(!ִ�п���ID)) & ",") = 0 Then
                    mstrִ�п���ID = mstrִ�п���ID & "," & Val(Nvl(!ִ�п���ID))
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
        Call InitCons
    End If
    
    mrsҽ��.Filter = 0
    If strFilter <> "" Then
        strFilter = Mid(strFilter, 5)
        mrsҽ��.Filter = strFilter
    End If
    With vsAdvice
            .Redraw = flexRDNone: .MergeCells = flexMergeNever
            '��ʱ�����ʱ��FormatString�ָ�һЩȱʡֵ(�̶����������̶��������ּ����ж���,�ߴ�,�ɼ�)
            'FormatString������ʱ��ֵ��Ч
            '���AutoResize=True,�������п���и߱��Զ�����(����AutoSizeMode)
            '���WordWrap=True,���и߻ᱻ�Զ�����
            .WordWrap = False
            Set .DataSource = mrsҽ��
            If mrsҽ��.RecordCount = 0 Then .Rows = .FixedRows + 1

            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            
           .WordWrap = True
            For i = 0 To .Cols - 1
                .ColKey(i) = Switch(i = 0, "ҽ��ͼ��", i = 1, "�����־", True, Trim(.TextMatrix(0, i)))
                .FixedAlignment(i) = flexAlignCenterCenter
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                Select Case .ColKey(i)
                Case "���", "�������", "��������", "�������", "��־", "�����", _
                        "ҽ��״̬", "ִ�б��", "���δ�ӡ", "ǩ����", "������", "����״̬", "���״̬", _
                        "�������"
                    .ColHidden(i) = True: .ColData(i) = "-1|1"
                Case "Ƥ��", "����", "����", "����", "Ƶ��"
                    .ColHidden(i) = True
                Case "ִ��ʱ��", "��ֹʱ��", "ִ������", "�ϴ�ִ��"
                    .ColHidden(i) = True
                Case "״̬", "����ʱ��", "У�Ի�ʿ", "У��ʱ��"
                    .ColHidden(i) = True
                Case "У��ʱ��", "ͣ��ʱ��", "ͣ����ʿ", "ȷ��ͣ��ʱ��"
                    .ColHidden(i) = True
                Case "���"
                    .ColData(i) = "1|0": .ColAlignment(i) = flexAlignCenterCenter
                    .ColDataType(i) = flexDTBoolean
                Case "ҽ������", "��Ч"
                    .ColData(i) = "1|0": .ColAlignment(i) = flexAlignLeftCenter
                Case Else
                    If .ColKey(i) Like "*ID" Then
                        .ColHidden(i) = True: .ColData(i) = "-1|1"
                    End If
                End Select
            Next
            For i = 1 To .Rows - 1
                .Cell(flexcpData, i, .ColIndex("Ӧ�ս��")) = .TextMatrix(i, .ColIndex("Ӧ�ս��"))
                .Cell(flexcpData, i, .ColIndex("ʵ�ս��")) = .TextMatrix(i, .ColIndex("ʵ�ս��"))
            Next
    End With
    If Not mrsҽ��.EOF Then
        Call ReModifyData(dt����)
    End If
    
    With vsAdvice
'            '������úϼ�
'            For i = 1 To .Rows - 1
'                lngҽ��ID = Val(.TextMatrix(i, .ColIndex("ҽ��ID")))
'                lng���ID = Val(.TextMatrix(i, .ColIndex("���ID")))
'                If lngҽ��ID = 0 Then lngҽ��ID = -1
'                If lng���ID = 0 Then lng���ID = -1
'
'            Next
                
        '�Զ������и�
        If InStr("2505,3345,1005,1335", .ColWidth(.ColIndex("�÷�"))) > 0 Then .ColWidth(.ColIndex("�÷�")) = IIf(mlngFontSize = 9, 2505, 3345)   '�û�δ�ĸ��п�ʱ������
        .AutoSize .ColIndex("����"), .ColIndex("�÷�")
        .ColWidth(.ColIndex("��ʼʱ��")) = IIf(mlngFontSize = 9, 1130, 1510)
        '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
        '����ǩ��ͼ�����
        .Cell(flexcpPictureAlignment, .FixedRows, .ColIndex("ҽ������"), .Rows - 1, .ColIndex("ҽ������")) = 0
        i = 0
         If lngҽ��ID <> 0 Then i = vsAdvice.FindRow(CStr(lngҽ��ID), , .ColIndex("ҽ��ID"))
        If i < .FixedRows Then i = .FixedRows
        .Row = i
        If .RowHidden(.Row) Then
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            For i = .Row - 1 To .FixedRows Step -1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
             .AddItem "":  .Row = .Rows - 1
        End If
        .Col = .FixedCols
        Call vsAdvice.ShowCell(.Row, .Col)
        zl_vsGrid_Para_Restore mlngModule, vsAdvice, Me.Caption, "ҽ�������ͷ��Ϣ"
        If mrsҽ��.RecordCount <> 0 And InStr(";" & mstrPrivs, ";��˲���;") > 0 Then
            vsAdvice.Editable = flexEDKbd
        Else
            vsAdvice.Editable = flexEDNone
        End If
        .Redraw = flexRDDirect
    End With
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    mblnChangeData = False
    Loadҽ�� = True
    Exit Function
ErrHand:
    vsAdvice.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitCons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ��ʼ������
    '����:���˺�
    '����:2012-05-31 16:11:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPreKey As String, rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    mblnNotClick = True
    With cboCons(CboIdx.EM_IDX�������)
        If .ListIndex >= 0 And .ListCount > 0 Then strPreKey = Chr(.ItemData(.ListIndex))
        .Clear
        .AddItem "�������"
        .ItemData(.NewIndex) = Asc("0")
        If mrs�������.RecordCount <> 0 Then mrs�������.MoveFirst
        Do While Not mrs�������.EOF
            If InStr(mstr������� & ",", "," & Nvl(mrs�������!����) & ",") > 0 Then
                .AddItem Nvl(mrs�������!����)
                .ItemData(.NewIndex) = Asc(Nvl(mrs�������!����))
                If strPreKey = Nvl(mrs�������!����) Then
                    .ListIndex = .NewIndex
                End If
            End If
            mrs�������.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
    End With
   If mstrִ�п���ID = "" Then Exit Sub
   strSQL = "" & _
    "   Select /*+ RULE */A.ID,A.����,A.����" & _
    "   From ���ű� A, (Select Column_Value From Table(f_num2list([1]))) J " & _
    "   Where A.ID=J. Column_Value " & _
    "   Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "0" & mstrִ�п���ID)
    
   With cboCons(CboIdx.EM_IDXִ�п���)
        If .ListIndex >= 0 And .ListCount > 0 Then strPreKey = .ItemData(.ListIndex)
        .Clear
        .AddItem "����ִ�п���"
        If InStr(1, mstrִ�п���ID & ",", ",-1,") > 0 Then
            .AddItem "<����>": .ItemData(.NewIndex) = -1
            If strPreKey = "-1" Then .ListIndex = .NewIndex
        End If
        If InStr(1, mstrִ�п���ID & ",", ",-2,") > 0 Then
            .AddItem "Ժ��ִ��": .ItemData(.NewIndex) = -2
            If strPreKey = "-2" Then .ListIndex = .NewIndex
        End If
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
                If strPreKey = Nvl(rsTemp!ID) Then
                    .ListIndex = .NewIndex
                End If
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub ReModifyData(ByVal dat���� As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '���:blnFilter-true,��ʾ����������
    '����:���˺�
    '����:2012-05-30 15:05:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��ҩ;��  As Boolean, bln��ҩ�÷�  As Boolean, bln�ɼ����� As Boolean, bln��Ѫ;��  As Boolean
    Dim i As Long, j As Long, strTemp  As String, strFormat As String
    Dim dtCurdate As Date, dtDate1 As Date, dtDate2 As Date
    Dim strCurDate As String, strDate1 As String, strDate2 As String
    Dim blnDo As Boolean, lngTop As Long, strTime As String, blnFirst As Boolean
    Dim strType As String '�������
    
    dtCurdate = zlDatabase.Currentdate: strCurDate = Format(dtCurdate, "yyyy-MM-DD")
    dtDate1 = DateAdd("D", -1, dtCurdate): strDate1 = Format(dtDate1, "yyyy-MM-DD")
    dtDate2 = DateAdd("D", -2, dtCurdate): strDate2 = Format(dtDate2, "yyyy-MM-DD")
    
    On Error GoTo errHandle
    With vsAdvice
        i = .FixedRows:
        Do While i <= .Rows - 1
            .Cell(flexcpData, i, .ColIndex("��ʼʱ��")) = CStr(.TextMatrix(i, .ColIndex("��ʼʱ��"))) '������ҩ�ӿڵ���ʱȡ��
            strTemp = Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "yyyy-MM-dd")
            Select Case strTemp
            Case strCurDate '����
                    .TextMatrix(i, .ColIndex("��ʼʱ��")) = "�� �� " & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "HH:mm")
            Case strDate1   '����
                    .TextMatrix(i, .ColIndex("��ʼʱ��")) = "�� �� " & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "HH:mm")
            Case strDate2  'ǰ��
                    .TextMatrix(i, .ColIndex("��ʼʱ��")) = "ǰ �� " & Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "HH:mm")
            Case Else
                    .TextMatrix(i, .ColIndex("��ʼʱ��")) = Format(.TextMatrix(i, .ColIndex("��ʼʱ��")), "MM-dd HH:mm")
            End Select
            'Ӧ�ս��,ʵ�ս��
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            
            bln��ҩ;�� = False: bln��ҩ�÷� = False: bln�ɼ����� = False: bln��Ѫ;�� = False
            If Trim(.TextMatrix(i, .ColIndex("�������"))) = "E" Then   '����
                     If Val(.TextMatrix(i - 1, .ColIndex("���ID"))) = Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) Then
                        Select Case .TextMatrix(i - 1, .ColIndex("�������"))
                        Case "5", "6"   '��ҩ���г�ҩ
                                bln��ҩ;�� = True
                                For j = i - 1 To .FixedRows Step -1
                                      If Val(.TextMatrix(j, .ColIndex("���ID"))) <> Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) Then Exit For
                                     '��ʾ��ҩ�ĸ�ҩ;��
                                    .TextMatrix(j, .ColIndex("�÷�")) = .TextMatrix(i, .ColIndex("�÷�"))
                                    '�ϲ��÷���:�÷� Ƶ�� ����
                                    strFormat = .TextMatrix(j, .ColIndex("�÷�"))
                                    strTemp = .TextMatrix(j, .ColIndex("Ƶ��"))
                                    If strTemp <> "" Then strFormat = strFormat & IIf(strFormat <> "", ",", "") & strTemp
                                    strTemp = .TextMatrix(j, .ColIndex("����"))
                                    If strTemp <> "" Then
                                        strFormat = strFormat & IIf(strFormat <> "", ",", "") & "��" & strTemp & "��"
                                    End If
                                     .TextMatrix(j, .ColIndex("�÷�")) = strFormat
                                     
                                     ''��ʾ��ҩ��ִ������
                                    If Val(.TextMatrix(j, .ColIndex("ִ������"))) = 5 And Val(.TextMatrix(i, .ColIndex("ִ������"))) <> 5 Then
                                        .TextMatrix(j, .ColIndex("ִ������")) = "�Ա�ҩ"
                                    ElseIf Val(.TextMatrix(j, .ColIndex("ִ������"))) <> 5 And Val(.TextMatrix(i, .ColIndex("ִ������"))) = 5 Then
                                        .TextMatrix(j, .ColIndex("ִ������")) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(j, .ColIndex("ִ������")) = IIf(Val(.TextMatrix(j, .ColIndex("ִ�б��"))) = 1, "��ȡҩ", "")
                                    End If
                                     .TextMatrix(j, .ColIndex("Ƥ��")) = .TextMatrix(i, .ColIndex("ҽ������"))
                                    If .TextMatrix(j, .ColIndex("Ƥ��")) <> "" Then
                                        .TextMatrix(j, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) & "," & .TextMatrix(j, .ColIndex("Ƥ��"))
                                    End If
                                Next
                            Case "7", "C" '�в�ҩ/����
                                    bln��ҩ�÷� = .TextMatrix(i - 1, .ColIndex("�������")) = "7" '��ҩ�÷���
                                    bln�ɼ����� = .TextMatrix(i - 1, .ColIndex("�������")) = "C" '�ɼ�������
                                    If bln�ɼ����� Then
                                        '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                                          j = .FindRow(.TextMatrix(i, .ColIndex("ҽ��ID")), .FixedRows, .ColIndex("���ID"))
                                        .TextMatrix(i, .ColIndex("�Թܱ���")) = .TextMatrix(j, .ColIndex("�Թܱ���"))
                                     End If
                                    '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                                    .TextMatrix(i, .ColIndex("ִ�п���")) = .TextMatrix(i - 1, .ColIndex("ִ�п���"))
                                     .TextMatrix(i, .ColIndex("ִ������")) = ""
                                     If bln��ҩ�÷� Then
                                        '��ʾ��ҩ�䷽ִ������
                                        If Val(.TextMatrix(i - 1, .ColIndex("ִ������"))) = 5 And Val(.TextMatrix(i, .ColIndex("ִ������"))) <> 5 Then
                                            .TextMatrix(i, .ColIndex("ִ������")) = "�Ա�ҩ"
                                        ElseIf Val(.TextMatrix(i - 1, .ColIndex("ִ������"))) <> 5 And Val(.TextMatrix(i, .ColIndex("ִ������"))) = 5 Then
                                            .TextMatrix(i, .ColIndex("ִ������")) = "��Ժ��ҩ"
                                        Else
                                            .TextMatrix(i, .ColIndex("ִ������")) = IIf(Val(.TextMatrix(i - 1, .ColIndex("ִ�б��"))) = 1, "��ȡҩ", "")
                                        End If
                                     End If
                                    'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ
                                    For j = i - 1 To .FixedRows Step -1
                                        If Val(.TextMatrix(j, .ColIndex("���ID"))) <> Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) Then Exit For
                                        .TextMatrix(i, .ColIndex("������")) = .TextMatrix(j, .ColIndex("������")) '���顢�䷽������ҽ��Ϊ׼
                                        .TextMatrix(i, .ColIndex("�ļ�ID")) = .TextMatrix(j, .ColIndex("�ļ�ID"))
                                        .RemoveItem j: i = i - 1
                                    Next
                            End Select
                     ElseIf .TextMatrix(i - 1, .ColIndex("�������")) = "K" And Val(.TextMatrix(i - 1, .ColIndex("ҽ��ID"))) = Val(.TextMatrix(i, .ColIndex("���ID"))) Then
                            bln��Ѫ;�� = True
                            '��ʾ��Ѫ;��
                            .TextMatrix(i - 1, .ColIndex("�÷�")) = .TextMatrix(i, .ColIndex("�÷�"))
                      Else
                         .TextMatrix(i, .ColIndex("ִ������")) = ""
                     End If
            End If
            
           '����ɼ��еĵ�һЩ��ʶ:�ſ����ɼ�����ʱδɾ������
            If Not (bln��ҩ;�� Or bln��Ѫ;��) And .TextMatrix(i, .ColIndex("�������")) <> "7" Then
                If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                '����ҽ���ָ�
                If Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) = 0 And .Rows > .FixedRows + 1 Then
                    .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = "�������� ����ҽ��(" & Format(dat����, "yyyy-MM-dd HH:mm") & ") ��������"
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
                    .Cell(flexcpAlignment, i, .FixedCols, i, .Cols - 1) = 4
                    .MergeRow(i) = True
                    .MergeCells = flexMergeFree
                End If
                If Left(.TextMatrix(i, .ColIndex("����")), 1) = "." Then
                    .TextMatrix(i, .ColIndex("����")) = "0" & .TextMatrix(i, .ColIndex("����"))
                End If
                If Left(.TextMatrix(i, .ColIndex("����")), 1) = "." Then
                    .TextMatrix(i, .ColIndex("����")) = "0" & .TextMatrix(i, .ColIndex("����"))
                End If
            
                If Val(.TextMatrix(i, .ColIndex("����ID"))) <> 0 Then
                        If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                            Set .Cell(flexcpPicture, i, .ColIndex("�����־")) = imgFlag.ListImages("����").Picture
                        ElseIf Val(.TextMatrix(i, .ColIndex("����״̬"))) = 1 Then
                            Set .Cell(flexcpPicture, i, .ColIndex("�����־")) = imgFlag.ListImages("��������").Picture
                        End If
                End If
                
                'ҽ����ɫ
                blnDo = False
                If Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 2 Then
                    'У������
                    If lngTop = 0 Then lngTop = i '��ɾ����Ҳ����Ӱ��ȡֵ
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80& '���
                    blnDo = True
                ElseIf Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 4 Then
                    '������
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                    .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                    blnDo = True
                ElseIf InStr(",8,9,", Val(.TextMatrix(i, .ColIndex("ҽ��״̬")))) > 0 Then
                    '��ֹͣ,��ȷ��ֹͣ:����������ֹʱ������ж�
                    If strCurDate >= .TextMatrix(i, .ColIndex("��ֹʱ��")) Or .TextMatrix(i, .ColIndex("��Ч")) = "����" Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 8 And strCurDate < .TextMatrix(i, .ColIndex("��ֹʱ��")) Then
                        '����,ֹͣ��,ֹͣʱ��δ����һ�����
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080 'ǳ��
                        blnDo = True
                    End If
                ElseIf Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 6 Then
                    '����ͣ
                    strTime = Format(GetAdviceTime(Val(.TextMatrix(i, .ColIndex("ҽ��ID"))), 6), "yyyy-MM-dd HH:mm")
                    If strCurDate >= strTime Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '����
                        blnDo = True
                    Else
                        '����,��ͣ��,��ͣʱ��δ����һ�����
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080 'ǳ��
                        blnDo = True
                    End If
                ElseIf Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 7 Then
                    '������
                    strTime = Format(GetAdviceTime(Val(.TextMatrix(i, .ColIndex("ҽ��ID"))), 7), "yyyy-MM-dd HH:mm")
                    If strCurDate < strTime Then
                        '����,���ú�,����ʱ��δ����һ�����
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H4AAD00 'ǳ��
                        blnDo = True
                    End If
                End If
                If Not blnDo Then
                    If lngTop = 0 Then lngTop = i
                    If Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) <> 1 And Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) <> 0 Then
                        '��ͨ��У��(Ҳ���������Ķ��״̬)
                        If Format(.TextMatrix(i, .ColIndex("�ϴ�ִ��")), "YYYY-MM-DD") >= Format(strCurDate, "YYYY-MM-DD") Then   '�����ѷ��͵�(�������ܷ��͵�����)
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HA08000               '����
                        Else
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '����
                        End If
                    End If
                End If
                'У�Ժ���ǰ����ҽ����ɫ��ʾ
                If .TextMatrix(i, .ColIndex("�������")) = "Z" And (Val(.TextMatrix(i, .ColIndex("��������"))) = 4 Or Val(.TextMatrix(i, .ColIndex("��������"))) = 14) _
                    And InStr(",-1,1,2,4,", Val(.TextMatrix(i, .ColIndex("ҽ��״̬")))) = 0 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '��ɫ
                End If
                
                '���ͺ�ת��ҽ����ɫ��ʾ
                If .TextMatrix(i, .ColIndex("�������")) = "Z" And Val(.TextMatrix(i, .ColIndex("��������"))) = 3 And Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) = 8 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '��ɫ
                End If
            
                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If .TextMatrix(i, .ColIndex("�������")) <> "" Then
                    If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(i, .ColIndex("�������"))) > 0 Then
                        .Cell(flexcpFontBold, i, .ColIndex("ҽ������")) = True
                    End If
                End If
                'Ƥ�Խ����ʶ
                If .TextMatrix(i, .ColIndex("�������")) = "E" And .TextMatrix(i, .ColIndex("��������")) = "1" And .TextMatrix(i, .ColIndex("Ƥ��")) <> "" Then
                    j = zl��ȡƤ�Խ��(Val(.TextMatrix(i, .ColIndex("������ĿID"))), .TextMatrix(i, .ColIndex("Ƥ��")))
                    .Cell(flexcpForeColor, i, .ColIndex("Ƥ��")) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, .ColIndex("Ƥ��")))
                End If
                '����¼��
                If Val(.TextMatrix(i, .ColIndex("������ĿID"))) = 0 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("����").Picture
                End If
                '������־:һ����ҩֻ��ʾ�ڵ�һ��
                blnFirst = True
                If InStr(",5,6,", .TextMatrix(i, .ColIndex("�������"))) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("���ID"))) = Val(.TextMatrix(i - 1, .ColIndex("���ID"))) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Val(.TextMatrix(i, .ColIndex("��־"))) = 1 Then
                        Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("����").Picture
                    ElseIf Val(.TextMatrix(i, .ColIndex("��־"))) = 2 Then
                        Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("��¼").Picture
                    End If
                    
                    If Val(.TextMatrix(i, .ColIndex("ҽ��״̬"))) < 2 Then   '�¿����ݴ��ҽ��
                        Select Case Val(.TextMatrix(i, .ColIndex("���״̬")))
                        '0-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ��
                            Case 1
                                Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("�����").Picture
                            Case 2
                                Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("���ͨ��").Picture
                            Case 3
                                Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("���δͨ��").Picture
                            Case Else
                        End Select
                        .Cell(flexcpPictureAlignment, i, .ColIndex("ҽ��ͼ��")) = 4
                    End If
                End If
                'δ��ҽ����ʶ
                If Val(.TextMatrix(i, .ColIndex("ִ�б��"))) = -1 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("δ��").Picture
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                End If
                
                'δ��ҽ����ʶ
                If Val(.TextMatrix(i, .ColIndex("ִ�б��"))) = -1 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("ҽ��ͼ��")) = imgFlag.ListImages("δ��").Picture
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                End If
                'Pass:�����������ʾ��ʾ��
                If .TextMatrix(i, .ColIndex("�����")) <> "" Then
                    Set .Cell(flexcpPicture, i, .ColIndex("�����")) = imgPass.ListImages(Val(.TextMatrix(i, .ColIndex("�����"))) + 1).Picture
                    .TextMatrix(i, .ColIndex("�����")) = ""
                End If
                '����ǩ����ʶ�����δ�ӡ��ʶ
                Call SetAdviceIcon(i)
            End If
            If bln��ҩ;�� Or bln��Ѫ;�� Then
                 .RemoveItem i
            Else
                '���ҽ������
                strFormat = .TextMatrix(i, .ColIndex("ҽ������"))
                If .TextMatrix(i, .ColIndex("�������")) <> "Z" And Val(.TextMatrix(i, .ColIndex("������ĿID"))) <> 0 And InStr(strFormat, "����ҽ��") = 0 Then
                    'ҽ�����ݶ����а����������ʱ�������ظ����
                    mrsDefine.Filter = "�������='" & .TextMatrix(i, .ColIndex("�������")) & "'"
                
                    strTemp = .TextMatrix(i, .ColIndex("Ƥ��"))
                    If strTemp <> "" Then strFormat = strFormat & strTemp
                    
                    If Not (InStr("5,6,7", .TextMatrix(i, .ColIndex("�������"))) = 0 And .TextMatrix(i, .ColIndex("Ƶ��")) = "һ����") Then
                        blnDo = True
                        If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                        If blnDo Then
                            strTemp = .TextMatrix(i, .ColIndex("����"))
                            If strTemp <> "" Then strFormat = strFormat & ",��" & strTemp
                        End If
                        
                        blnDo = True
                        If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                        If blnDo Then
                            strTemp = .TextMatrix(i, .ColIndex("����"))
                            If strTemp <> "" Then strFormat = strFormat & ",ÿ��" & strTemp
                        End If
                    End If
                End If
                
                .TextMatrix(i, .ColIndex("����")) = strFormat
                '�ϲ��÷���:�÷� Ƶ�� ����(һ����ҩ����ǰ���Ѵ���)
                If .TextMatrix(i, .ColIndex("�������")) <> "Z" And Val(.TextMatrix(i, .ColIndex("������ĿID"))) <> 0 And InStr(strFormat, "����ҽ��") = 0 Then
                    strFormat = .TextMatrix(i, .ColIndex("�÷�"))
                    strTemp = .TextMatrix(i, .ColIndex("Ƶ��"))
                    If strTemp <> "" Then strFormat = strFormat & IIf(strFormat <> "", ",", "") & strTemp
                    
                    strTemp = .TextMatrix(i, .ColIndex("����"))
                    If strTemp <> "" Then
                        strFormat = strFormat & IIf(strFormat <> "", ",", "") & "��" & strTemp & "��"
                    End If
                    .TextMatrix(i, .ColIndex("�÷�")) = strFormat
                End If
                i = i + 1
            End If
        Loop
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ModiyStartDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ�ʼʱ��
    '����:���˺�
    '����:2012-05-30 15:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Sub
Private Sub InitAdvice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������
    '����:���˺�
    '����:2012-05-30 14:54:18
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsAdvice
        .Clear 1
        .Rows = vsAdvice.FixedRows + 1
        .Editable = flexEDNone
        For i = .FixedRows To .Rows - 1
            .RowHidden(i) = False
        Next
    End With
End Sub
Private Function GetAdviceTime(ByVal lngҽ��ID As Long, ByVal int���� As Integer) As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ָ��������ʱ��
    '����:���˺�
    '����:2012-05-30 16:44:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(����ʱ��) as ʱ�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, int����)
    If rsTemp.EOF Then Exit Function
    If Not IsNull(rsTemp!ʱ��) Then GetAdviceTime = rsTemp!ʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zl��ȡƤ�Խ��(ByVal lng��Ŀid As Long, ByVal str��� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƥ�Խ����ע������������
    '���:str���=Ƥ�Խ����ע����,��"(+)"
    '����:-1-����,1-����,0-�޽��
    '����:���˺�
    '����:2012-05-30 16:50:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim var���� As Variant, var���� As Variant
    Dim strSQL As String, i As Integer
    On Error GoTo errH
    If mrsSkinTest Is Nothing Then
        strSQL = "Select ID,Nvl(�걾��λ,'����(+);����(-)') as ��ע From ������ĿĿ¼ Where ���='E' And ��������='1'"
        Set mrsSkinTest = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    mrsSkinTest.Filter = "ID=" & lng��Ŀid
    If mrsSkinTest.EOF Then Exit Function
    var���� = Split(Split(mrsSkinTest!��ע, ";")(0), ",")
    var���� = Split(Split(mrsSkinTest!��ע, ";")(1), ",")
    
    For i = 0 To UBound(var����)
        If Right(var����(i), Len(str���)) = str��� Then
            zl��ȡƤ�Խ�� = 1: Exit Function
        End If
    Next
    For i = 0 To UBound(var����)
        If Right(var����(i), Len(str���)) = str��� Then
            zl��ȡƤ�Խ�� = -1: Exit Function
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    '(518716, 1)
    Call InitData
    Call InitPancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mrsSkinTest Is Nothing Then
        If mrsSkinTest.State <> 1 Then mrsSkinTest.Close
    End If
    Set mrsSkinTest = Nothing
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "ҽ�������ͷ��Ϣ"
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
End Sub

Private Sub SetAdviceIcon(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�е���������ҽ�����ݵ�ͼ���ʶ
    '����:���˺�
    '����:2012-05-30 17:02:56
    '˵��:ע���ǵ������ã�����һ������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsAdvice
        If Val(.TextMatrix(lngRow, .ColIndex("ǩ����"))) = 1 And Val(.TextMatrix(lngRow, .ColIndex("���δ�ӡ"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("ҽ������")) = img16dbl.ListImages(1).Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("����")) = img16dbl.ListImages(1).Picture
        ElseIf Val(.TextMatrix(lngRow, .ColIndex("ǩ����"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("ҽ������")) = img16.ListImages("ǩ��").Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("����")) = img16.ListImages("ǩ��").Picture
        ElseIf Val(.TextMatrix(lngRow, .ColIndex("���δ�ӡ"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("ҽ������")) = img16.ListImages("���δ�ӡ").Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("����")) = img16.ListImages("���δ�ӡ").Picture
        Else
            Set .Cell(flexcpPicture, lngRow, .ColIndex("ҽ������")) = Nothing
            Set .Cell(flexcpPicture, lngRow, .ColIndex("����")) = Nothing
        End If
    End With
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2012-05-30 17:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    mlngFontSize = 9: mlngModule = glngModul: mblnChangeData = False
    strSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Set mrsDefine = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    strSQL = "Select ����,���� From ������Ŀ���"
    Set mrs������� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mstrPrivs = gstrPrivs
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckPatiDataMoved(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����˵������Ƿ���ת��
    '����:���˺�
    '����:2012-05-30 17:30:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    strSQL = "Select ����ת�� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ת��", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!����ת��) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

 
Private Sub imgColList_Click(Index As Integer)
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgList(Index).hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList(Index).Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, IIf(Index = 1, vsAdvice, vsFeeList), lngLeft, lngTop, imgColList(Index).Height)
    If Index = 1 Then
        zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "ҽ�������ͷ��Ϣ"
    Else
        zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
    End If
End Sub

Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    cboCons(0).Top = (picFilter.ScaleHeight - cboCons(0).Height) \ 2
    cboCons(1).Top = cboCons(0).Top
    lblCons(0).Top = cboCons(0).Top + (cboCons(0).Height - lblCons(0).Height) \ 2
    lblCons(1).Top = lblCons(0).Top
    lblCons(2).Top = lblCons(0).Top
    chkType(0).Top = cboCons(0).Top + (cboCons(0).Height - chkType(0).Height) \ 2
    chkType(1).Top = chkType(0).Top
    chk����.Top = chkType(0).Top
    
    cboCons(0).Left = lblCons(0).Left + lblCons(0).Width + 10
    lblCons(1).Left = cboCons(0).Left + cboCons(0).Width + 50
    cboCons(1).Left = lblCons(1).Left + lblCons(1).Width + 10
    lblCons(2).Left = cboCons(1).Left + cboCons(1).Width + 50
    chkType(0).Left = lblCons(2).Left + lblCons(2).Width + 10
    chkType(1).Left = chkType(0).Left + chkType(0).Width + 20
    chk����.Left = chkType(1).Left + chkType(1).Width + 50
    
End Sub

Private Sub picImgList_Click(Index As Integer)
    Call imgColList_Click(Index)
End Sub
Private Sub picҽ��_Resize()
    Err = 0: On Error Resume Next
    With picҽ��
        
        vsAdvice.Left = .ScaleLeft + 10
        vsAdvice.Top = .ScaleTop
        vsAdvice.Height = .ScaleHeight - 20
        vsAdvice.Width = .ScaleWidth
        picImgList(1).Top = vsAdvice.Top + 30
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft + 10
        vsFeeList.Top = .ScaleTop
        vsFeeList.Height = .ScaleHeight
        vsFeeList.Width = .ScaleWidth - 20
    End With
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngҽ��ID As Long, bln��� As Boolean
    With vsAdvice
        If Row <= 0 Then Exit Sub
        If Col <> .ColIndex("���") Then Exit Sub
        lngҽ��ID = Val(.TextMatrix(Row, .ColIndex("ҽ��ID")))
        If lngҽ��ID = 0 Then Exit Sub
        bln��� = Val(.TextMatrix(Row, Col)) <> 0
        bln��� = SaveData(lngҽ��ID, bln���)
        If bln��� = False Then
            .TextMatrix(Row, Col) = IIf(bln���, 0, 1)
        End If
    End With
End Sub

Private Sub vsAdvice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngҽ��ID As Long, lngCol As Long
    Dim lng���ID As Long
    
    If NewRow = OldRow And vsAdvice.Visible = False Then Exit Sub
    On Error GoTo errHandle
    With vsAdvice
        lngCol = .ColIndex("ҽ��ID")
        If NewRow > 0 And lngCol >= 0 Then lngҽ��ID = Val(.TextMatrix(NewRow, lngCol))
        lngCol = .ColIndex("���ID")
        If NewRow > 0 And lngCol >= 0 Then lng���ID = Val(.TextMatrix(NewRow, lngCol))
    End With
    Call LoadFeeList(lngҽ��ID, lng���ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    With vsAdvice
        Select Case Col
        Case .ColIndex("ҽ������"), .ColIndex("����")
                .AutoSize Col, .ColIndex("�÷�")
        Case .ColIndex("Ƥ��")
            If .ColWidth(Col) > 1200 Then .ColWidth(Col) = 1200
        Case Else
            If Row = -1 Then
                lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
                If vsAdvice.ColWidth(Col) < lngW Then
                    vsAdvice.ColWidth(Col) = lngW
                ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
                    vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
                End If
            End If
        End Select
    End With
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "ҽ�������ͷ��Ϣ"
End Sub



Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        Select Case Col
        Case .ColIndex("���")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then Cancel = True: Exit Sub
        Select Case Col
        Case .ColIndex("�����"), .ColIndex("���")
             Cancel = True: Exit Sub
        Case Else
        End Select
    End With
    If Row = -1 Then
        With vsAdvice
            If Col <= .FixedCols - 1 Then
                Cancel = True
            ElseIf Col = .ColIndex("�����") Then
                Cancel = True
            End If
        End With
    End If
End Sub
Private Sub LoadFeeList(ByVal lngҽ��ID As Long, Optional lng���ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�����ϸ
    '����:���˺�
    '����:2012-05-30 17:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblӦ�� As Double, dblʵ�� As Double
    Dim i As Long
    On Error GoTo errHandle
    strSQL = _
        " Select a.No, a.�۸񸸺�, a.���, a.�շ�ϸĿid, a.ִ�в���id, a.��¼״̬, a.ִ��״̬," & _
        "        a.�Ǽ�ʱ��, a.����, a.����, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.ҽ�����" & _
        " From סԺ���ü�¼ A" & _
        " Where a.����id= [1] And (a.��ҳid = [2] Or a.��ҳid Is Null)"
    strSQL = strSQL & " Union All " & Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    
    strSQL = _
    " Select b.No as ���ݺ�, b.���, b.�շ�ϸĿid,B.��¼״̬,Q.���� As �շ�����," & _
    "        q.���, b.����, Decode(r.סԺ��λ, Null, q.���㵥λ, r.סԺ��λ) As ��λ, " & _
    "        to_char(b.���� * Nvl(r.סԺ��װ, 1),'999999999999990.99') As ����, " & _
    "        b.Ӧ�ս��, b.ʵ�ս��, j.���� As ִ�п���, " & _
    "        Decode(b.��¼״̬, 2, -1 * b.ִ��״̬ || '���˷�', Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 1, '��ȫִ��', 2, '����ִ��', '�쳣�շ�')) As ִ��״̬, " & _
    "        To_Char(b.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ�� " & _
    " From (Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, a.ִ�в���id, a.��¼״̬, a.ִ��״̬, a. �Ǽ�ʱ��, " & _
    "              Avg(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����, Avg(a.��׼����) As ����," & _
    "              Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��" & _
    "       From (" & strSQL & ") A,����ҽ����¼ B" & _
    "       Where A.ҽ�����=B.ID And (B.ID in ([3],[4]) or nvl(B.���ID,-2) in ([3],[4] )) " & _
    "       Group By a.No, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, a.ִ�в���id, a.��¼״̬, a.ִ��״̬, a.�Ǽ�ʱ��) B," & _
    "      ���ű� J, �շ���ĿĿ¼ Q, ҩƷ��� R " & _
    " Where b.ִ�в���id = j.Id(+) And b.�շ�ϸĿid = q.Id And b.�շ�ϸĿid = r.ҩƷid(+) " & _
    " Order By �Ǽ�ʱ��, ���ݺ�,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, lngҽ��ID, lng���ID)
    With vsFeeList
        .Clear 1
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 0 To .Cols - 1
            .ColKey(i) = IIf(i = 0, "�̶���־", .TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "��¼״̬" Then
                .ColHidden(i) = True:  .ColData(i) = "-1|1"
            End If
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "�̶���־"
                    .ColData(i) = "-1|1"
            Case "���ݺ�", "��λ"
                .ColData(i) = "1|0"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "�շ�����"
                .ColData(i) = "1|0"
            Case "����", "Ӧ�ս��", "ʵ�ս��", "����"
                .ColData(i) = "1|0"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
        .ColWidth(.ColIndex("�̶���־")) = 300
        dblӦ�� = 0: dblʵ�� = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(.TextMatrix(i, .ColIndex("����"))), 5)
            .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), "######" & gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), "######" & gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), "######" & gstrDec)
            dblӦ�� = dblӦ�� + Val(.TextMatrix(i, .ColIndex("Ӧ�ս��")))
            dblʵ�� = dblʵ�� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            Select Case Val(.TextMatrix(i, .ColIndex("��¼״̬")))
            Case 2
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
            Case 3
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbBlue
            Case Else
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
            End Select
        Next
        If rsTemp.RecordCount <> 0 Then
            .Rows = .Rows + 1: i = .Rows - 1
            .TextMatrix(i, .ColIndex("���ݺ�")) = "�ϼ�"
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(dblӦ��, "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(dblʵ��, "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    '      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
    '      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)
            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '����һ����ҩ������еı��߼�����
            lngLeft = vsAdvice.ColIndex("��Ч"): lngRight = vsAdvice.ColIndex("��ʼʱ��")
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = vsAdvice.ColIndex("����"): lngRight = vsAdvice.ColIndex("�÷�")
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = vsAdvice.ColIndex("Ƥ��"): lngRight = vsAdvice.ColIndex("Ƥ��")
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub
Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, .ColIndex("�������")) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, .ColIndex("�������"))) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function


Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, strPrompt As String
    strPrompt = ""
    With vsAdvice
            lngRow = vsAdvice.MouseRow
            If Not (Button = 0 And lngRow > 0) Then Exit Sub
            Select Case .MouseCol
            Case .ColIndex("����")
            Case .ColIndex("ҽ��ͼ��")
                If Val(.TextMatrix(lngRow, .ColIndex("������ĿID"))) = 0 Then
                    strPrompt = "����¼���ҽ��"
                ElseIf Val(.TextMatrix(lngRow, .ColIndex("��־"))) = 1 Then
                    strPrompt = "����ҽ��"
                ElseIf Val(.TextMatrix(lngRow, .ColIndex("��־"))) = 2 Then
                    strPrompt = "��¼ҽ��"
                End If
                 '����п�����ҩ�����Ϣ��������ʾ
                If Val(.TextMatrix(lngRow, .ColIndex("ҽ��״̬"))) = 1 Then
                    Select Case Val(.TextMatrix(lngRow, .ColIndex("���״̬")))
                    Case 1
                        strPrompt = "������ҩ�����"
                    Case 2
                        strPrompt = "������ҩ���ͨ��"
                    Case 3
                       strPrompt = "������ҩ���δͨ��:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, .ColIndex("ҽ��ID"))))
                    End Select
                End If
            End Select
            If strPrompt <> "" Then
               Call zlCommFun.ShowTipInfo(vsAdvice.hWnd, strPrompt)
            Else
                Call zlCommFun.ShowTipInfo(vsAdvice.hWnd, "")
            End If
    End With
End Sub

Private Function GetKSSAuditQuestion(ByVal lngҽ��ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ҩ���δͨ���ķ�����Ϣ
    '����:���˺�
    '����:2012-05-31 14:40:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(����˵��,'��') as ����˵�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=12 Order by ����ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID)
    If Not rsTmp.EOF Then GetKSSAuditQuestion = rsTmp!����˵��
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "���������ͷ��Ϣ"
End Sub

Private Sub vsFeeList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsFeeList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long

    With vsFeeList
        If Col > .FixedCols - 1 Then Done = True: Exit Sub
    '�����̶����еı����
    SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)
    '����߱����
    vRect.Left = Left
    vRect.Top = Top
    vRect.Right = Left + 1
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '���ϱ߱����
    vRect.Left = Left
    vRect.Top = Top
    vRect.Right = Right
    vRect.Bottom = Top + 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '���±߱����
    vRect.Left = Left
    vRect.Top = Bottom - 1
    vRect.Right = Right
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '���ұ߱����
    vRect.Left = Right - 1
    vRect.Top = Top
    vRect.Right = Right
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    End With
    Done = True
End Sub
Private Function SaveData(ByVal lngҽ��ID As Long, ByVal bln��� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2012-05-31 17:58:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    'Zl_����ҽ����¼_�������
    strSQL = "Zl_����ҽ����¼_�������("
    'Id_In           ����ҽ����¼.Id%Type,
    strSQL = strSQL & "" & lngҽ��ID & ","
    '�Ƿ�������_In ����ҽ����¼.�Ƿ�������%Type
    strSQL = strSQL & "" & IIf(bln���, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mblnChangeData = True
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

