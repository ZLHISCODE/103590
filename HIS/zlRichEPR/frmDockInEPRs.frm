VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInEPRs 
   BorderStyle     =   0  'None
   Caption         =   "סԺ������¼"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   1335
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   6255
      _cx             =   11033
      _cy             =   2355
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   135
      ScaleHeight     =   3120
      ScaleWidth      =   8145
      TabIndex        =   0
      Top             =   195
      Width           =   8145
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockInEPRs.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3480
         Left            =   735
         TabIndex        =   1
         Top             =   165
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   6138
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockInEPRs.frx":054E
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         Editable        =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   3105
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   8070
         _cx             =   14235
         _cy             =   5477
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   1
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6855
            Picture         =   "frmDockInEPRs.frx":059C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   4
            Top             =   255
            Width           =   250
         End
         Begin MSComctlLib.ImageList imgThis 
            Left            =   0
            Top             =   1125
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":6DEE
                  Key             =   "��д"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7388
                  Key             =   "�޶�"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7922
                  Key             =   "�鵵"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7EBC
                  Key             =   "ת��"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":8256
                  Key             =   "��ӡ"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   90
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   15
      Top             =   705
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDockInEPRs.frx":EAB8
      Left            =   720
      Top             =   4785
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------
'���峣��
'-----------------------------------------------------
Private Enum mCol
    ��־ = 0: ���˿���: ҳ������: ��������: ������: ����ʱ��: ������: ���ʱ��: ��ǰ�汾: ǩ������: ��ǰ���: �鵵��: �鵵����: ����ID: ������: ����: ����״̬:  ����: ID: ��������: ҳ����: �༭��ʽ: ��ӡ: Ӥ��: �걨״̬: ������¼
End Enum

Const conDefColWidth = "270;0;1200;1600;800;1600;800;1600;500;0;3300;0;0;0;1200;0;0;0;0;0;0;0;0;0;0;0"
Const conPane_List = 1
Const conPane_Content = 2
Const conPane_New = 3
Const mlngModul = 1251
Private mstrColWidthConfig As String
Private mlngfolding As Long
'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '�̳��ĵ�����ġ������ϲο��¼���
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ���߶Ա�����(1250)��Ȩ�޴�
Private mblnSearch As Boolean   '��ǰʹ�����Ƿ�߱���������(1273)Ȩ
Private mlngPatiId As Long      '����id
Private mlngPageId As Long      '��ҳid
Private mlngDeptId As Long      '��ǰ��������id����һ���ǵ�ǰ���˿���
Private mblnEdit As Boolean     '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˿��Ҿ�����
Private mblnMoved As Boolean    '�Ƿ������Ѿ�ת��
Private mlngAdviceID As Long    'ҽ��ID
Private mintState As Integer    '��clsDockInEPR
Private mblnInsideTools As Boolean '�Խ�������
Private mstrPhysicians  As String '��������ҽʦ���ִ�
Private mblnAllowDelete As Boolean '�Ƿ�����ɾ��
Private mblnShowFinal As Boolean '��ʾ���հ汾
Private mlngCurId As Long

Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockInContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Private mfrmTipInfo As New frmTipInfo

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR            '���ʽ�����༭��
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mcbsThis As Object          'CommandBar�ؼ�
Private mlngVersion As Long         'ѡ�е��ļ��汾��
Private mblnDisease As Boolean      '�Ƿ�ӵ����1249ģ���Ȩ��


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If cbrMain.Count < 1 Then Exit Sub
    zlUpdateCommandBars Control
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane.ID
    Case conPane_New
        Select Case Action
        Case PaneActionClosing, PaneActionClosed: Cancel = False
        Case Else: Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_List
            Item.Handle = picList.hwnd
        Case conPane_Content
            Item.Handle = mfrmContent.hwnd
        Case conPane_New
            Item.Handle = mfrmNew.hwnd
    End Select
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    Call Event_AfterPrinted(lngRecordId)
End Sub
Public Sub Event_AfterPrinted(lngRecordId As Long)
Dim i As Integer
    For i = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(i, mCol.ID) = lngRecordId Then
            vfgThis.Cell(flexcpData, i, mCol.��ǰ���) = ""
            vfgThis.Cell(flexcpText, i, mCol.��ӡ) = gstrUserName
            Set vfgThis.Cell(flexcpPicture, i, mCol.ҳ������) = imgThis.ListImages("��ӡ").Picture
            Exit For
        End If
    Next
End Sub
Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��ʾָ�������б��е���ʷǩ����¼
Dim strTipInfo As String, lngRow As Long, strPrint As String
    If picInfo.Visible = False Then Exit Sub
    lngRow = vfgThis.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vfgThis.Cell(flexcpData, lngRow, mCol.��ǰ���)
    
    If strTipInfo = "" Then '���û�л�ȡ������������ȡ����¼���б���
        strTipInfo = GetEprSign(vfgThis.TextMatrix(lngRow, mCol.ID))   '��ȡǩ��
        Call EprPrinted(vfgThis.TextMatrix(lngRow, mCol.ID), strPrint) '��ȡ��ӡ��¼
        strTipInfo = "�� " & Rpad(vfgThis.TextMatrix(lngRow, mCol.������), 8) & _
                     "�� " & Rpad(vfgThis.TextMatrix(lngRow, mCol.����ʱ��), 19) & " ����" & vbCrLf & strTipInfo
        strTipInfo = strTipInfo & vbCrLf & strPrint
        vfgThis.Cell(flexcpData, lngRow, mCol.��ǰ���) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hwnd, strTipInfo, True
End Sub
Private Sub piclist_Resize()
On Error Resume Next
    With vfgThis
        .Top = 0: .Left = 0
        .Width = picList.ScaleWidth: .Height = picList.ScaleHeight
    End With

    fraColSel.Move Me.vfgThis.Left + 50, Me.vfgThis.Top + 50
    fraColSel.ZOrder 0
    vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
    vsColumn.ZOrder 0
Err.Clear
End Sub

Private Sub vfgThis_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
If picInfo.Visible Then
    picInfo.Move vfgThis.Cell(flexcpLeft, NewTopRow, mCol.��ǰ���) + vfgThis.Cell(flexcpWidth, NewTopRow, mCol.��ǰ���) - picInfo.Width - 30
End If
End Sub

Private Sub vfgThis_Click()
    Dim lngMouseRow As Long, lngMouseCol As Long, lngWidth As Long, i As Long
    With vfgThis
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                If DisplayContent(Val(.TextMatrix(lngMouseRow, mCol.ID))) Then
                    With vsfFeedback
                        .Left = vfgThis.Left + vfgThis.Width - .Width
                        .Top = vfgThis.Top + 300 * (lngMouseRow + 1)
                        .ZOrder
                        .Visible = True
                        .SetFocus
                    End With
                End If
            Else
                vsfFeedback.Visible = False
            End If
        End If
    End With
End Sub

Private Sub vfgThis_KeyDown(KeyCode As Integer, Shift As Integer)
    vsColumn_KeyDown KeyCode, Shift
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Long
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                vfgThis.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vfgThis.ColHidden(.RowData(i)) Or vfgThis.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
        
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '��ѡ����
    Else
        If Me.vfgThis.Visible Then Me.vfgThis.SetFocus
    End If
    RaiseEvent Activate
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    vsColumn.Visible = False '��ѡ����
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vfgThis.MouseRow = -1 And Me.Tag = "" Then
        vfgThis.Row = vfgThis.Rows - 1
    End If
End Sub

Private Sub vfgThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, lngRow As Long
    lngCol = vfgThis.MouseCol: lngRow = vfgThis.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl.Name <> "vfgThis" Then
            vfgThis.SetFocus
        Else
            vfgThis.SetFocus
        End If
    Else
        vfgThis.SetFocus
    End If
    
    If Val(vfgThis.TextMatrix(lngRow, mCol.ID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vfgThis.Cell(flexcpLeft, lngRow, mCol.��ǰ���) + vfgThis.Cell(flexcpWidth, lngRow, mCol.��ǰ���) - picInfo.Width - 30, vfgThis.Cell(flexcpTop, lngRow, mCol.��ǰ���) + 15
        If vfgThis.RowSel = lngRow Then
            picInfo.BackColor = vfgThis.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
    If lngRow >= 0 And lngRow < vfgThis.Rows And lngCol >= 0 And lngCol < vfgThis.Cols Then
        If vfgThis.Cell(flexcpFontUnderline, lngRow, lngCol) = True Then
            vfgThis.MousePointer = 54
        Else
            vfgThis.MousePointer = 0
            If vsfFeedback.Visible Then vsfFeedback.Visible = False
        End If
    Else
        If vsfFeedback.Visible Then vsfFeedback.Visible = False
    End If
End Sub

Private Sub vfgThis_SelChange()
    If picInfo.Visible Then
        picInfo.BackColor = vfgThis.BackColorSel
    End If
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split(conDefColWidth, ";")
            vfgThis.ColWidth(lngCol) = T(lngCol)
            vfgThis.ColHidden(lngCol) = False
        Else
            vfgThis.ColWidth(lngCol) = 0
            vfgThis.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To 19
        strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
    Next
    mstrColWidthConfig = strCols
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '�ر���ѡ����
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vfgThis.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '����ѡ����
        Call imgColSel_MouseUp(1, 0, 0, 0)
    End If
End Sub

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String, i As Long
    If vfgThis.Cols = UBound(Split(conDefColWidth, ";")) + 1 Then
        For i = 0 To vfgThis.Cols - 1
            If i = mCol.������¼ Then vfgThis.ColWidth(i) = 0
            strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
        Next
    Else
        strCols = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", conDefColWidth)
    End If
    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", vfgThis.FontSize
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "ShowHistory", IIf(mblnShowFinal, "True", "False")
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    If Not mfrmPrintPreview Is Nothing Then Unload mfrmPrintPreview
    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mobjDoc = Nothing
    Set mfrmPrintPreview = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mfrmTipInfo = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub Form_Load()
Dim panList As Pane, panContent As Pane, panNew As Pane, lngFontSize As Long
    mlngPatiId = -1: mlngPageId = -1
    mblnShowFinal = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "ShowHistory", "True")
    mstrColWidthConfig = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", conDefColWidth)
    lngFontSize = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", 9)
    vfgThis.FontSize = lngFontSize
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)
    mlngfolding = zlDatabase.GetPara("�������۵���ʼ����", glngSys, mlngModul, "6")
    mstrPrivs = GetPrivFunc(glngSys, 1251)
    
    Set panList = dkpMan.CreatePane(conPane_List, 200, 300, DockTopOf, Nothing)
    panList.Title = "�����б�"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmContent = New frmDockInContent
    Set panContent = dkpMan.CreatePane(conPane_Content, 200, 300, DockBottomOf, Nothing)
    panContent.Title = "��������"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmNew = New frmDockEPRNew
    Set panNew = dkpMan.CreatePane(conPane_New, 100, 400, DockRightOf, Nothing)
    panNew.Title = "��������"
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    Set mObjTabEprView = New cTableEPR
    Call mObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDBUser)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With Me.cbrMain
'        .VisualTheme = xtpThemeOfficeXP
        .VisualTheme = xtpThemeOffice2003
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Visible = False
        .EnableCustomization False
        With .Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '����VisualTheme����Ч
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
        End With
    End With
    
    Me.dkpMan.SetCommandBars Me.cbrMain
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    mlngVersion = 1  'Ĭ��Ϊ��1��
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
    Dim rs As New ADODB.Recordset, rt As RECT
    Dim strFileName As String, blnResult As Boolean
    Dim frmThis As Form, bFinded As Boolean
    
    If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, mlngModul, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
    End If
    
    If UserNewEMR Then
        MsgBox "�������Ѿ���ʼʹ���²���ϵͳ����ʹ���²���ϵͳ��д������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If gstrPrivsEpr = ";;" Then
        MsgBox "�����߱������༭��ӦȨ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnMoved Then
        MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errHand

    If Split(EprIsCommit, "|")(0) = 0 Then
        MsgBox "�ò��˲������ύ��飬����������������ȡ���������ԣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TimeLimitOut Then Exit Sub
        
    gstrSQL = "Select ����,���� From �����ļ��б� Where  ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
    If rs!���� < 0 Then
        '���ⲡ������������
        Exit Sub
    ElseIf rs!���� = 2 Then '���ʽ�༭��
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_����, cprET_�������༭, FileId, True, 0, cprPF_סԺ, mlngPatiId, mlngPageId, babyNum, mlngDeptId, mlngAdviceID, mstrPrivs, , InStr(gstrPrivsEpr, "������ӡ") > 0, Val(gstrESign)
        End If
    ElseIf rs!���� = 4 Then '��Ⱦ�����濨�༭��
'        ��Ⱦ���Ѷ���ҳ��
    Else                    '���Ӳ���RichEpr
        If InStr(rs!����, "������¼") > 0 Or InStr(rs!����, "��������") > 0 Or InStr(rs!����, "��������¼") > 0 Then '��Ҫ��ҽ�����к�ѡ
            gstrSQL = "Select a.Id, b.���� ҽ��, c.���� ִ�п���, a.����ҽ��, To_Char(a.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') ��ʼʱ��," & vbNewLine & _
                        "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��" & vbNewLine & _
                        "From ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                        "Where a.����id = [1] And a.��ҳid = [2] and a.���ID IS NULL And a.������Ŀid = b.Id And b.��� = 'Z' And b.�������� = '7' And a.ִ�п���id = c.Id(+) And" & vbNewLine & _
                        "      Not Exists (Select 1 From ����ҽ������ C Where c.ҽ��id = a.Id)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ӧҽ��", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�����Ӧҽ��", False, 1, "�����Ӧҽ�������ڶ��ҽ���Ļ����¼�ɶԳ���", False, False, False, 0, 0, 0, blnResult, True, True, mlngPatiId, mlngPageId)
                If blnResult = True Then 'ȡ��ѡ��
                    MsgBox "����ҽ����д�����¼����Ҫָ������ҽ����", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    mlngAdviceID = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                mlngAdviceID = rs!ID
            Else '�����ݣ�δ������ҽ�������ѿ�����ҽ���Ѿ���д ������¼ �������� ��������¼;����ҽԺ��Ҫ���´����ҽ����������д�����¼����ͨ��
                'MsgBox "��δ�¿�����ҽ�������Ѿ���д����ҽ����ز��������飡", vbExclamation, gstrSysName: Exit Sub
            End If
        ElseIf InStr(rs!����, "����") > 0 And InStr(rs!����, "��") = 0 Then '��Ҫ�� ������¼���������롢��������¼ ��ѡ
            gstrSQL = "Select b.ҽ��id As ID, f.���� ҽ��ִ�п���, a.��������, a.������, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��" & vbNewLine & _
                        "From ���Ӳ�����¼ A, ����ҽ������ B, ����ҽ����¼ E, ���ű� F" & vbNewLine & _
                        "Where a.����id = [1] And a.��ҳid = [2] And a.�������� = 2 And a.Id = b.����id And" & vbNewLine & _
                        "      (Instr(a.��������, '������¼') > 0 Or Instr(a.��������, '��������') > 0 Or Instr(a.��������, '��������¼') > 0) And b.ҽ��id = e.Id And" & vbNewLine & _
                        "      e.���id Is Null And e.ִ�п���id = f.Id(+) And Not Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From ���Ӳ�����¼ C, ����ҽ������ D" & vbNewLine & _
                        "       Where d.ҽ��id = b.ҽ��id And d.����id = c.Id And c.�������� = 2 And Instr(c.��������, '����') > 0 And Instr(c.��������, '��') = 0)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�����¼��Ӧҽ��", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�����¼��Ӧ����", False, 1, "�����Ӧ���������ڶ�������¼ʱ�ɶԳ���", False, False, False, 0, 0, 0, blnResult, True, True, mlngPatiId, mlngPageId)
                If blnResult = True Then 'ȡ��ѡ��
                    MsgBox "����ҽ����д�����¼����Ҫָ���������룡", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    mlngAdviceID = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                mlngAdviceID = rs!ID
            Else '������ �����Ǹ���ǰ��д��������¼��ҽ��ID �� ������¼�Ѿ���д�����¼
                
            End If
        Else
            If zlDatabase.GetPara("��������������д��������", glngSys, mlngModul, "1") = 1 Then
                '�жϹ����ĵ��Ƿ��Ѿ���д��
                gstrSQL = "Select ID From �����ļ��б� Where ��� <> NVL(ҳ��,���) And ID =[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
                If rs.EOF = False Then '�ǹ����ĵ�
                    gstrSQL = "Select M.ID,M.����" & vbNewLine & _
                                "       From �����ļ��б� L, �����ļ��б� M" & vbNewLine & _
                                "       Where M.���� = L.���� And M.��� = L.ҳ�� And L.ID =[1]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
                    If rs.EOF Then MsgBox "�ò����Ĺ���������ʧЧ������ϵϵͳ����Ա��", vbInformation, gstrSysName: Exit Sub
                    strFileName = rs!����
                    gstrSQL = "Select ID" & vbNewLine & _
                                "From ���Ӳ�����¼" & vbNewLine & _
                                "Where ����id = [1] And ��ҳid =[2] And �ļ�id+0 =[3]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, Val(rs!ID))
                    If rs.EOF Then
                        MsgBox "�ò����Ĺ����� [" & strFileName & "] ��δ��д�����顣", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.����ID = mlngPatiId _
                        And .EPRPatiRecInfo.������Դ = cprPF_סԺ And .EPRPatiRecInfo.��ҳID = mlngPageId _
                        And .EPRPatiRecInfo.����ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                        Exit For
                    End If
                End With
            End If
        Next
        
        If bFinded = False Then
            Set mobjDoc = New cEPRDocument
            mobjDoc.InitEPRDoc cprEM_����, cprET_�������༭, FileId, cprPF_סԺ, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
            mobjDoc.EPRPatiRecInfo.Ӥ�� = babyNum
            mobjDoc.ShowEPREditor Me
        End If
    End If
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjDoc_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton And Not mcbsThis Is Nothing Then
        Dim Popup As CommandBar
        Dim ControlBar As CommandBarControl
        
        Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&U)"): ControlBar.BeginGroup = True
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵(&I)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Sort, "��������(&S)"): ControlBar.BeginGroup = True
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "����ҽ����������(&C)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim byteEdit As Byte
    Dim ControlBar As Object
    On Error GoTo errHand
    Me.dkpMan.Panes(conPane_New).Close
    With Me.vfgThis
        If .Rows <= 1 Then Exit Sub
        If .Cols < mCol.ID + 1 Then Exit Sub
        mlngCurId = Val(.TextMatrix(.Row, mCol.ID))
        byteEdit = Val(.TextMatrix(.Row, mCol.�༭��ʽ))
    End With
    If Not mcbsThis Is Nothing Then
        Set ControlBar = mcbsThis.FindControl(, conMenu_Edit_Delete, , True)
        zlUpdateCommandBars ControlBar
        If Not mcbsThis.FindControl(, conMenu_Edit_Delete, , True) Is Nothing Then
            mblnAllowDelete = mcbsThis.FindControl(, conMenu_Edit_Delete, , True).Enabled
        End If
    End If
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> mlngCurId) Then
        Me.Tag = "Refresh" '����ˢ��̫�죬���򱨡��ܾ�Ȩ�ޡ�
        Call mfrmContent.zlRefresh(mlngCurId, IIf(mblnEdit = False, "", mstrPrivs), mblnMoved, mblnShowFinal, byteEdit, mblnAllowDelete)
        Me.Tag = ""
        Me.vfgThis.Tag = mlngCurId
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zlDefCommandBars(ByVal cbsThis As Object, ByVal blnInsideTools As Boolean)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

    mblnInsideTools = blnInsideTools
    Set mcbsThis = cbsThis
    Set mcbsThis.Icons = zlCommFun.GetPubIcons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportAll, "��������RTF�ļ�(&A)��", cbrControl.Index + 1): cbrControl.ToolTipText = "�����ò�������ȫ��ʽ����ΪRTF"
        '���ڵ���ΪXML�ļ�֮��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "�б��ӡ(&T)", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "��������(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
        Set cbrControl = .Add(xtpControlButton, ID_PATISIGNVerify, "����ǩ����֤(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "ҽ������")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowHistory, "��ʾ���հ汾(&L)", cbrControl.Index + 1)
        cbrControl.BeginGroup = True
    End With

    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "�����������(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "���˲�������(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "�������İ�(&C)")
    End With
    
    '����������
    '-----------------------------------------------------
    cbrMain.DeleteAll
    If mblnInsideTools Then
        Set cbrToolBar = cbrMain.Add("������", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ContextMenuPresent = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ"): cbrControl.STYLE = xtpButtonIconAndCaption
        End With
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��", 1)
            .Item(cbrControl.Index + 1).BeginGroup = True
        End With
    End If
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options
        .AddHiddenCommand conMenu_Edit_Archive
        .AddHiddenCommand conMenu_Edit_Untread
    End With
    
    '-----------------------------------------------------
    '����Ȩ��״̬����ʾ���Ӵ���
    '-----------------------------------------------------
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0) Then
        If zlDatabase.GetPara("�Զ���ʾ�������", glngSys, mlngModul, "1") = 1 Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String
Dim rs As New ADODB.Recordset
Dim strSQL As String, lFileId As Long, blnCanPrint As Boolean
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_Audit Or Control.ID = conMenu_Edit_Archive Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML Or conMenu_Edit_Compend) Then '��ת������,�޸�,ɾ��,���,�鵵,�򿪲��������
        MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If

    lFileId = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    bEditor = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.�༭��ʽ))
    Select Case Control.ID
    Case conMenu_File_Open
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        blnCanPrint = InStr(1, gstrPrivsEpr, "������ӡ") > 0
        If blnCanPrint Then blnCanPrint = (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.���ʱ��)) <> "" Or InStr(1, gstrPrivsEpr, "δǩ����ӡ") > 0)
        If blnCanPrint Then blnCanPrint = (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
        If blnCanPrint Then blnCanPrint = IIf(EprPrinted(lFileId), InStr(mstrPrivs, "ȡ����ӡ") > 0, True) '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
        If blnCanPrint Then blnCanPrint = (vfgThis.TextMatrix(vfgThis.Row, mCol.������) = gstrUserName Or InStr(1, mstrPrivs, "��������") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0) '������д���в�������Ȩ��,��������ҽʦ
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView '�鿴�ò���
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint, , mlngAdviceID
        ElseIf bEditor = 1 Then
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, True, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, blnCanPrint, Val(gstrESign)
            End If
        ElseIf bEditor = 2 Then
'            ��Ⱦ���Ѷ���ҳ��
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
            MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(True)
    Case conMenu_File_Print
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
            MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(False)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_ExportAll: Call ExportAll
    Case conMenu_File_ExportToXML
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        '������XML�ļ�
        Dim strF As String
        dlgThis.Filename = "����_" & vfgThis.TextMatrix(vfgThis.Row, mCol.��������) & "(" & vfgThis.TextMatrix(vfgThis.Row, mCol.ID) & "," & mlngVersion & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        With Me.vfgThis
            If Val(.TextMatrix(.Row, mCol.��������)) = 2 And Val(.TextMatrix(.Row, mCol.����)) < 0 Then
                '�����סԺ����
            ElseIf bEditor = 1 Then
                '���ʽ����
                mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, False, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved
                If mObjTabEprView.zlExportXML(strF) Then
                    MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
                End If
            Else
                '��ͨסԺ����
                Dim DocXML As New cEPRDocument
                DocXML.InitAndOpenEPR lFileId, mlngVersion, , True
                If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                    DoEvents
                    MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
                End If
            End If
        End With
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    Case conMenu_Edit_NewItem
        Me.dkpMan.Panes(conPane_New).Select
        Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
    Case conMenu_Edit_Modify
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If TimeLimitOut Then Exit Sub '������¼ʱ�ޣ��������޸ģ����������
        '�������༭ģʽ
        With Me.vfgThis
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
            If Val(.TextMatrix(.Row, mCol.��������)) = 2 And Val(.TextMatrix(.Row, mCol.����)) < 0 Then
                '�����סԺ�������������¼��
            ElseIf bEditor = 1 Then
                '���ʽ����
                If Not mObjTabEpr Is Nothing Then
                    bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
                End If
                If bFinded = False Then
                    Set mObjTabEpr = New cTableEPR
                    mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_סԺ, _
                        mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(1, gstrPrivsEpr, "������ӡ") > 0, Val(gstrESign)
                End If
            ElseIf bEditor = 0 Then
                'RichEPR����
                For Each frmThis In Forms
                    If frmThis.Name = "frmMain" Then
                        On Error Resume Next
                        If frmThis.Document.EPRPatiRecInfo.ID = .TextMatrix(.Row, mCol.ID) And frmThis.Document.EPRPatiRecInfo.����ID = mlngPatiId _
                            And frmThis.Document.EPRPatiRecInfo.������Դ = cprPF_סԺ And frmThis.Document.EPRPatiRecInfo.��ҳID = mlngPageId _
                            And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            bFinded = True
                        End If
                    End If
                Next
                If bFinded = False Then
                    Set mobjDoc = New cEPRDocument
                    mobjDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, .TextMatrix(.Row, mCol.ID), cprPF_סԺ, mlngPatiId, CStr(mlngPageId), 0, mlngDeptId, mlngAdviceID
                    mobjDoc.ShowEPREditor Me, InStr(1, gstrPrivsEpr, "������ӡ") > 0
                End If
            ElseIf bEditor = 2 Then
'                ��Ⱦ���Ѷ�����ʾ
            End If
        End With
    Case conMenu_Edit_Delete
        If Split(EprIsCommit, "|")(1) = 0 Then
            MsgBox "�ò��˲������ύ��飬����ɾ������ȡ���������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    
        With Me.vfgThis
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
            strInfo = "���ɾ����ݡ�" & .TextMatrix(.Row, mCol.��������) & "����"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_���Ӳ�����¼_Delete(" & .TextMatrix(.Row, mCol.ID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End With
    Case conMenu_Edit_Audit
        If GetCurrentGdi > 8000 Then Call MsgBox("��ǰϵͳ��Դռ�ù��࣬���ȹر�һЩ�����༭���ں������ԣ�", vbInformation, gstrSysName): Exit Sub
        
        If TimeLimitOut Then Exit Sub '������¼ʱ�ޣ��������޸ģ����������
        If EprPrinted(lFileId) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
        
        If bEditor = 1 Then
            '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_סԺ, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, True, 0, cprPF_סԺ, _
                    mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, , Val(gstrESign)
            End If
        Else
            '���������ģʽ
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    On Error Resume Next
                    If frmAudit.Document.EPRPatiRecInfo.ID = lFileId _
                        And frmAudit.Document.EPRPatiRecInfo.������Դ = cprPF_סԺ And frmAudit.Document.EPRPatiRecInfo.����ID = mlngPatiId _
                        And frmAudit.Document.EPRPatiRecInfo.��ҳID = mlngPageId And frmAudit.ChildMode = False Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_�޸�, cprET_���������, Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.ID), cprPF_סԺ, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
                mobjDoc.ShowEPREditor Me, InStr(1, gstrPrivsEpr, "������ӡ") > 0
            End If
        End If
    Case conMenu_Edit_Archive
        Call EprArchive
    Case conMenu_Edit_Sort
        '����
        Dim frmSort As New frmEPRSort
        If frmSort.ShowMe(Me, mlngPatiId, mlngPageId, vfgThis.TextMatrix(vfgThis.Row, mCol.��������), vfgThis.TextMatrix(vfgThis.Row, mCol.ҳ����)) = True Then
            'ˢ����ʾ
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End If
    Case conMenu_Edit_NoPrint 'ȡ����ӡ���
        If Split(EprIsCommit, "|")(0) = 0 Then
            MsgBox "�ò��˲������ύ��飬���ܳ�����ӡ����ȡ���������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        Call PrintCancel(CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)))
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 2, mlngDeptId, 1, mintState)
    Case conMenu_Tool_Search: Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_View_Refresh: Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    Case conMenu_Edit_Compend
        Call modelsApply
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '���ʽ������28δ��������ǩ�����
            'call
        End If
    Case ID_PATISIGNVerify
        Call VerifyPatiSign(Me, lFileId, mblnMoved)
    Case conMenu_View_ShowHistory
        mblnShowFinal = Not mblnShowFinal
        vfgThis.Tag = 0: Call RefreshList
    Case conMenu_Edit_ApplyModi
        Err = 0: On Error GoTo errHand
        Dim lngOrderId As Long
            gstrSQL = "Select a.Id, b.���� ҽ��, c.���� ִ�п���, a.����ҽ��, To_Char(a.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') ��ʼʱ��," & vbNewLine & _
                        "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��" & vbNewLine & _
                        "From ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                        "Where a.����id = [1] And a.��ҳid = [2] and a.���ID IS NULL And a.������Ŀid = b.Id And b.��� = 'Z' And b.�������� = '7' And a.ִ�п���id = c.Id(+)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ӧҽ��", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�����Ӧҽ��", False, 1, "�����Ӧҽ�������ڶ��ҽ���Ļ����¼�ɶԳ���", False, False, False, 0, 0, 0, bFinded, True, True, mlngPatiId, mlngPageId)
                If bFinded = True Then 'ȡ��ѡ��
                    MsgBox "����ҽ����д�����¼����Ҫָ������ҽ����", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    lngOrderId = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                lngOrderId = rs!ID
            Else '�����ݣ�δ������ҽ�������ѿ�����ҽ���Ѿ���д ������¼ �������� ��������¼;����ҽԺ��Ҫ���´����ҽ����������д�����¼����ͨ��
                'MsgBox "��δ�¿�����ҽ�������Ѿ���д����ҽ����ز��������飡", vbExclamation, gstrSysName:
                Exit Sub
            End If
        gstrSQL = "Zl_����ҽ������_Modify(" & lFileId & "," & lngOrderId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum
    Dim blnTmp As Boolean
    
    With Me.vfgThis
        Select Case Control.ID
        Case conMenu_File_Open, conMenu_File_Excel, conMenu_File_RowPrint
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_NoPrint
            Control.Enabled = InStr(mstrPrivs, "ȡ����ӡ") > 0 And (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            If Control.Enabled Then Control.Enabled = Trim(.TextMatrix(.Row, mCol.��ӡ)) <> ""
            If Control.Enabled Then Control.Enabled = mblnEdit
        Case conMenu_Edit_NewItem
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "������ӡ") > 0)
            If Control.Enabled Then Control.Enabled = IIf(Trim(.TextMatrix(.Row, mCol.���ʱ��)) = "", InStr(1, gstrPrivsEpr, "δǩ����ӡ") > 0, True)
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
            If Control.Enabled Then Control.Enabled = (vfgThis.TextMatrix(vfgThis.Row, mCol.������) = gstrUserName Or InStr(1, mstrPrivs, "��������") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '������д���в�������Ȩ��,��������ҽʦ
            If Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_ExportToXML Then
                If Control.Enabled Then Control.Enabled = Val(.TextMatrix(.Row, mCol.�༭��ʽ)) <> 2
            End If
        Case conMenu_File_ExportAll
            Control.Enabled = (Val(.TextMatrix(1, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "������ӡ") > 0)
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '������д���в�������Ȩ��,��������ҽʦ
        Case conMenu_Edit_Modify
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
            If Control.Enabled And Not mblnDisease Then
                blnTmp = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)  '�Ѿ������������Ĳ������ܴ���
                If Not blnTmp Then
                    If Val(.TextMatrix(.Row, mCol.�걨״̬)) = 4 Or Val(.TextMatrix(.Row, mCol.�걨״̬)) = 5 Then
                        blnTmp = True
                    End If
                End If
                Control.Enabled = blnTmp
            End If
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.����ID)))   '���Ʋ����ſ��Ը�
            If Control.Enabled Then
                If Trim(.TextMatrix(.Row, mCol.���ʱ��)) = "" Then
                    Control.Enabled = (InStr(1, mstrPrivs, "���˲���") > 0 Or Trim(.TextMatrix(.Row, mCol.������)) = Trim(gstrUserName))
                ElseIf Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" And Val(.TextMatrix(.Row, mCol.��ǰ�汾)) <= 1 And InStr(1, ",1,2,4,", Val(.TextMatrix(.Row, mCol.ǩ������))) > 0 Then
                    Control.Enabled = (InStr(1, mstrPrivs, "���˲���") > 0 Or InStr(1, .TextMatrix(.Row, mCol.������), Trim(gstrUserName)) > 0)
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Delete
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0) And (mblnEdit And mlngPatiId > 0 And (InStr(1, mstrPrivs, "������д") > 0 Or InStr(1, mstrPrivs, "ǿ��ɾ��") > 0))
            If Control.Enabled And InStr(1, mstrPrivs, "ǿ��ɾ��") > 0 Then Exit Sub '�߱�ǿ��ɾ��Ȩ�ޣ��򲻽��к������ж�
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)  '�Ѿ������������Ĳ������ܴ���
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.����ID)))   '���Ʋ����ſ���ɾ
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.���ʱ��)) = "")        'δ��ɲ�������ɾ
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "���˲���") > 0 Or Trim(.TextMatrix(.Row, mCol.������)) = Trim(gstrUserName))
        Case conMenu_Edit_Audit
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "��������") > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)  '�Ѿ������������Ĳ������ܴ���
'            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.����ID)))   '���Ʋ����ſ������
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.��������)) <> 2 Or Val(.TextMatrix(.Row, mCol.����)) >= 0) '�����סԺ�������������¼�����ṩ����
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.���ʱ��)) <> "")       '��ɲ����ſ�����
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.�鵵��)) = "")          'δ�鵵����������
            If Control.Enabled Then Control.Enabled = Val(.TextMatrix(.Row, mCol.�༭��ʽ)) <> 2           '��Ⱦ�����濨����֧���޶�
        Case conMenu_Edit_Archive
            Control.Enabled = (mblnEdit And mlngPatiId > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)  '�Ѿ������������Ĳ������ܴ���
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.ǩ������)) <> 0)         '��ǰ�汾�Ѿ�ǩ����ɲſ��Թ鵵
            If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
                Control.Caption = "�鵵": Control.Checked = False
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����鵵") > 0)
            Else
                Control.Caption = "����": Control.Checked = True
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������") > 0)
            End If
        Case conMenu_Edit_Sort
            '����ֻ�ж��ĵ�����ҳ��ʱ�ſ��Ե�����ţ�
            Control.Visible = True: Control.Enabled = True
            Control.Visible = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_Compend
            Control.Enabled = InStr(1, mstrPrivs, "������д") > 0
        Case conMenu_Tool_Monitor
            Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "�������") > 0)
        Case conMenu_Tool_Search: Control.Enabled = mblnSearch
        Case conMenu_Tool_SignVerify
            Control.Enabled = Val(.TextMatrix(.Row, mCol.ID)) <> 0 And Trim(.TextMatrix(.Row, mCol.���ʱ��)) <> ""
        Case conMenu_View_ShowHistory
            Control.Checked = mblnShowFinal
        Case conMenu_Edit_ApplyModi
            Control.Visible = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            Control.Visible = InStr(.TextMatrix(.Row, mCol.��������), "����") > 0
            Control.Enabled = Trim(.TextMatrix(.Row, mCol.��ӡ)) = ""
        End Select
    End With
End Sub

Public Sub RefreshList()
    Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
End Sub

Private Sub InitColumnSelect()
    On Error Resume Next
    '���ܣ�����ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vfgThis
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.��������, mCol.������, mCol.����ʱ��, mCol.������, mCol.���ʱ��, mCol.��ǰ�汾, mCol.��ǰ���, mCol.������, mCol.Ӥ��
                 vsColumn.Rows = vsColumn.Rows + 1
                 lngRow = vsColumn.Rows - 1
                 vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                 vsColumn.RowData(lngRow) = i
                
                 '�̶���ʾ��
                 If InStr(",ҳ������,��������,", "," & .TextMatrix(0, i) & ",") > 0 Then
                     vsColumn.TextMatrix(lngRow, 0) = 1
                     vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                 End If
            End Select
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-С(ȱʡ)��1-��
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub

Private Sub Initvfg()
    With vfgThis
        On Error Resume Next
        mfrmContent.Clear
        .Tag = ""
        .Clear
        .Rows = 1
        .Cols = 26
        .TextMatrix(0, mCol.��־) = "��־"
        .TextMatrix(0, mCol.���˿���) = "���˿���"
        .TextMatrix(0, mCol.ҳ������) = "��������"
        .TextMatrix(0, mCol.��������) = "��������"
        .TextMatrix(0, mCol.������) = "������"
        .TextMatrix(0, mCol.����ʱ��) = "����ʱ��"
        .TextMatrix(0, mCol.������) = "������"
        .TextMatrix(0, mCol.���ʱ��) = "���ʱ��"
        .TextMatrix(0, mCol.��ǰ�汾) = "�汾"
        .TextMatrix(0, mCol.ǩ������) = "ǩ������"
        .TextMatrix(0, mCol.��ǰ���) = "��ǰ���"
        .TextMatrix(0, mCol.�鵵��) = "�鵵��"
        .TextMatrix(0, mCol.�鵵����) = "�鵵����"
        .TextMatrix(0, mCol.����ID) = "����ID"
        .TextMatrix(0, mCol.������) = "������"
        .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.����״̬) = "����״̬"
        .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.ID) = "ID"
        .TextMatrix(0, mCol.��������) = "��������"
        .TextMatrix(0, mCol.ҳ����) = "ҳ����"
        .TextMatrix(0, mCol.�༭��ʽ) = "�༭��ʽ"
        .TextMatrix(0, mCol.��ӡ) = "��ӡ"
        .TextMatrix(0, mCol.�걨״̬) = "�걨״̬"
        .TextMatrix(0, mCol.Ӥ��) = "Ӥ��"
        .TextMatrix(0, mCol.������¼) = "������¼"
        .MergeCellsFixed = flexMergeFree
        .MergeCol(mCol.ҳ������) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        
        Dim T As Variant, i As Long '�����п�
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) <> .Cols - 1 Then
            mstrColWidthConfig = conDefColWidth
            T = Split(mstrColWidthConfig, ";")
        End If
        For i = 0 To .Cols - 1
            .ColWidth(i) = T(i)
            .ColHidden(i) = (.ColWidth(i) = 0)
        Next
        
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = mCol.ҳ������
        .SubtotalPosition = flexSTAbove
    End With
    
    vsfFeedback.Visible = False
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long _
                            , Optional ByVal intState As Integer) As Long
    Dim lngCurId As Long    '��ǰ������¼ID
    Dim lngCurRow As Long   'ˢ�º�ѡ���кţ�Ĭ��Ϊ0����ѡ��
    Dim rsTemp As New ADODB.Recordset, rsDis As ADODB.Recordset
    Dim lngCol As Long, lngRow As Long, i As Long
    Dim strKind As String, blnGroupTurnDept As Boolean
    Dim strReportIDs As String
    Dim str��Ⱦ������ As String
    Dim rs��Ⱦ As ADODB.Recordset
    Dim str���� As String
    
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function
    lngCurId = IIf(mlngPatiId = lngPatiID, Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)), 0) '��ǰ����ˢ��ǰѡ����ID
    If lngCurId = 0 Then lngCurId = mlngCurId
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '��ȡ�Ƿ񱾲������õ���ǩ��,���ұ����ûȡ��ʱ��ȡ
        gstrESign = getPassESign(1, lngDeptId)
    End If
    
    mblnDisease = (GetPrivFunc(glngSys, 1249) <> "")   'true-�����˼�������ģ��;false-�����ü�������ģ��
    
    mlngDeptId = lngDeptId
    mblnEdit = blnEdit
    mblnMoved = blnMoved
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mlngAdviceID = lngAdviceID
    mintState = intState
    vsColumn.Visible = False
    mstrPhysicians = GetPhysicians '��ȡ����ҽ������
    blnGroupTurnDept = (zlDatabase.GetPara("ת�ƺ�Ҫ����д�Ĺ���������һҳ��ӡ", glngSys, mlngModul, 1) = 1)
    picInfo.Visible = False
    Call Initvfg
    
    If mblnDisease Then
        str���� = "r.�������� In (2, 6)"
    Else
        str���� = " r.�������� In (2, 5, 6) "
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select r.����id ���˿���, Decode(b.����, Null, r.��������, b.����) As ҳ��, r.��������, r.������ As ������," & vbNewLine & _
                "       To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, r.������, To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��," & vbNewLine & _
                "       r.���汾 As ��ǰ�汾, r.ǩ������," & vbNewLine & _
                "       Decode(r.���汾, 1, '��д��', '�޶���') || r.������ || '��' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') ||" & vbNewLine & _
                "        Decode(Nvl(r.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') As ��ǰ���, r.�鵵��, r.�鵵����, r.����id, d.���� As ������, c.����, r.����״̬," & vbNewLine & _
                "       Decode(c.���, b.���, 1, 0) As ����, r.Id, r.��������, b.���, r.�༭��ʽ, r.��ӡ�� As ��ӡ, r.Ӥ��, e.ҽ��id" & vbNewLine & _
                "From ���Ӳ�����¼ R, ���ű� D, �����ļ��б� C, ����ҳ���ʽ B, ����ҽ������ E" & vbNewLine & _
                "Where r.�ļ�id + 0 = c.Id And r.������Դ = 2 And " & str���� & " And r.����id = d.Id And r.����id = [1] And r.��ҳid = [2] And" & vbNewLine & _
                "      r.Id = e.����id(+) And c.���� = b.���� And c.ҳ�� = b.��� And Nvl(c.����, 0) <> 4" & vbNewLine & _
                "Order By r.��������, b.���, e.ҽ��id, r.���, r.����ʱ��"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    If Not mblnDisease Then
        gstrSQL = "Select a.����״̬,b.id From �����걨��¼ a,���Ӳ�����¼ b, ������ҳ c  where a.�ļ�id=b.id and b.��������=5" & vbNewLine & _
            "and b.����id+0=c.����id and b.��ҳid+0=c.��ҳid and a.����=c.���� and c.����id=[1] and c.��ҳid=[2] and a.����״̬ in (4,5)"
        Set rs��Ⱦ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        
        For lngRow = 1 To rs��Ⱦ.RecordCount
            str��Ⱦ������ = str��Ⱦ������ & "," & rs��Ⱦ!ID
            rs��Ⱦ.MoveNext
        Next
    End If
    
    strKind = ""
    With vfgThis
        .ColWidth(mCol.�걨״̬) = 0
        .ColHidden(mCol.�걨״̬) = True
        .ColWidth(mCol.������¼) = 0
        .ColHidden(mCol.������¼) = True
        Do Until rsTemp.EOF
            .Rows = .Rows + 1
            .IsSubtotal(.Rows - 1) = True

            .TextMatrix(rsTemp.AbsolutePosition, mCol.���˿���) = NVL(rsTemp!���˿���)
            .Cell(flexcpData, rsTemp.AbsolutePosition, mCol.ҳ������) = NVL(rsTemp!ҳ��)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ҳ������) = NVL(rsTemp!ҳ��)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.��������) = NVL(rsTemp!��������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.������) = NVL(rsTemp!������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����ʱ��) = NVL(rsTemp!����ʱ��)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.������) = NVL(rsTemp!������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.���ʱ��) = NVL(rsTemp!���ʱ��)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.��ǰ�汾) = NVL(rsTemp!��ǰ�汾)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ǩ������) = NVL(rsTemp!ǩ������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.��ǰ���) = NVL(rsTemp!��ǰ���)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.�鵵��) = NVL(rsTemp!�鵵��)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.�鵵����) = NVL(rsTemp!�鵵����)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����ID) = NVL(rsTemp!����ID)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.������) = NVL(rsTemp!������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����) = NVL(rsTemp!����)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����״̬) = NVL(rsTemp!����״̬)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����) = NVL(rsTemp!����)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ID) = NVL(rsTemp!ID)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.��������) = NVL(rsTemp!��������)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ҳ����) = NVL(rsTemp!���)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.�༭��ʽ) = NVL(rsTemp!�༭��ʽ)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.��ӡ) = NVL(rsTemp!��ӡ)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.Ӥ��) = NVL(rsTemp!Ӥ��)
            If str��Ⱦ������ <> "" Then
                If InStr(str��Ⱦ������ & ",", "," & rsTemp!ID & ",") > 0 Then
                    rs��Ⱦ.Filter = "id=" & rsTemp!ID
                    If Not rs��Ⱦ.EOF Then
                        .TextMatrix(rsTemp.AbsolutePosition, mCol.�걨״̬) = Val(rs��Ⱦ!����״̬ & "")
                        .ColWidth(mCol.������¼) = 1200
                        .ColHidden(mCol.������¼) = False
                        .TextMatrix(rsTemp.AbsolutePosition, mCol.������¼) = "������¼"
                        .Cell(flexcpForeColor, rsTemp.AbsolutePosition, mCol.������¼, rsTemp.AbsolutePosition, mCol.������¼) = &HFF0000     '��ɫ
                        .Cell(flexcpFontUnderline, rsTemp.AbsolutePosition, mCol.������¼, rsTemp.AbsolutePosition, mCol.������¼) = True
                        End If
                End If
            End If

            'ҳ��������ͬ���飬����ʱ�����飬ת��ʱ������
            If .Cell(flexcpData, rsTemp.AbsolutePosition - 1, mCol.ҳ������) = NVL(rsTemp!ҳ��) And NVL(rsTemp!����, 0) <> 1 _
                And Not (blnGroupTurnDept And .TextMatrix(rsTemp.AbsolutePosition - 1, mCol.���˿���) <> NVL(rsTemp!���˿���)) Then
                .RowOutlineLevel(rsTemp.AbsolutePosition) = 1
                .TextMatrix(rsTemp.AbsolutePosition, mCol.ҳ������) = ""
            Else
                .RowOutlineLevel(rsTemp.AbsolutePosition) = 0
            End If
            
            If strKind <> .TextMatrix(rsTemp.AbsolutePosition, mCol.��������) Then '����������
                If strKind <> "" Then Call .CellBorderRange(rsTemp.AbsolutePosition, 0, rsTemp.AbsolutePosition, .Cols - 1, RGB(0, 0, 255), 0, 1, 0, 0, 0, 0)
                strKind = .TextMatrix(rsTemp.AbsolutePosition, mCol.��������)
            End If

            If Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.����״̬)) > 0 Then '״̬ͼ��
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��־) = imgThis.ListImages("ת��").Picture
            ElseIf Trim(.TextMatrix(rsTemp.AbsolutePosition, mCol.�鵵��)) <> "" Then
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��־) = imgThis.ListImages("�鵵").Picture
            ElseIf Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.��ǰ�汾)) <= 1 Then
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��־) = imgThis.ListImages("��д").Picture
            Else
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��־) = imgThis.ListImages("�޶�").Picture
            End If
            .MergeRow(rsTemp.AbsolutePosition) = True
            If Trim(.TextMatrix(rsTemp.AbsolutePosition, mCol.��ӡ)) <> "" Then '��ӡͼ��
                 If NVL(rsTemp!ҳ��) <> NVL(rsTemp!��������) Or .RowOutlineLevel(rsTemp.AbsolutePosition) = 1 Then
                    .Cell(flexcpPictureAlignment, rsTemp.AbsolutePosition, mCol.��������) = flexAlignLeftCenter
                    Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��������) = imgThis.ListImages("��ӡ").Picture
                Else
                    .Cell(flexcpPictureAlignment, rsTemp.AbsolutePosition, mCol.ҳ������) = flexAlignLeftCenter
                    Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.ҳ������) = imgThis.ListImages("��ӡ").Picture
                End If
            Else
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.ҳ������) = Nothing
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.��������) = Nothing
            End If
            
            If .ROWHEIGHT(rsTemp.AbsolutePosition) < .RowHeightMin Then .ROWHEIGHT(rsTemp.AbsolutePosition) = .RowHeightMin
            If lngCurId = Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.ID)) Then lngCurRow = rsTemp.AbsolutePosition '��ֵ�к�
            rsTemp.MoveNext
        Loop
        
        Call Folding '�����۵�
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngCurRow = 0 Then
            vfgThis.Tag = -1: .Row = 0 '��ʹvfgthis��ѡ���κ��У�����ʾ�κ����ݣ�����ѡ��ĳ��ʱ��ˢ��
        Else
           .Row = lngCurRow
        End If
        Call vfgThis_RowColChange
        zlRefresh = .Rows - 1
    End With
    Call InitColumnSelect '��ѡ����
    
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0) Then '����Ȩ��״̬����ʾ���Ӵ���
        If zlDatabase.GetPara("�Զ���ʾ�������", glngSys, mlngModul, "1") = 1 Then
            dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
    Exit Function
   
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Folding()
Dim i As Long, l As Long, N As Long
    l = CLng(vfgThis.Height / vfgThis.RowHeightMin)
    
    If vfgThis.Rows > l Then '����������С��ʵ������
        For i = 1 To vfgThis.Rows - 1
            If vfgThis.RowOutlineLevel(i) = 1 Then '��������,��ʼ����������6��ʱ
                N = N + 1
            Else
                N = 0
            End If
            
            If N >= mlngfolding Then
                vfgThis.IsCollapsed(i - mlngfolding) = flexOutlineCollapsed: N = 0
            End If
        Next i
    End If
End Sub

Private Sub AutoResizeCol(ByVal intCol As Integer)
    Dim intRow As Integer
    Dim lngMaxWidth As Long
    
    With vfgThis
        For intRow = .FixedRows To .Rows - 1
            If lngMaxWidth < LenB(.TextMatrix(intRow, intCol)) Then
                lngMaxWidth = LenB(.TextMatrix(intRow, intCol))
            End If
        Next

        If lngMaxWidth > 0 Then
            .ColWidth(intCol) = lngMaxWidth * 90 + 120
        End If
    
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '       strSubhead����ӡ�ĸ�����
    '-------------------------------------------------
Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
Dim rsTemp As New ADODB.Recordset
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = "������д���"
    
    '---------------------------------------------
    '��û�����Ϣ
    Dim strSubhead As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select a.סԺ��, a.���� From ������Ϣ a Where a.����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId)
    If Not rsTemp.EOF Then
        strSubhead = "סԺ��:" & rsTemp!סԺ�� & "  ����:" & rsTemp!����
    Else
        strSubhead = ""
    End If
    Err = 0: On Error GoTo 0
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead)
    Call objAppRow.Add("��" & mlngPageId & "��סԺ")
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'## ���ܣ�  ��ʽ����Ԥ������ӡ
'##
'## ������  blnPreview  :�Ƿ���Ԥ��ģʽ
'################################################################################################################
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long, strPrintName As String
Dim r As String, blnOrigMode As Boolean  '�Ƿ���ʾԭʼ״̬
    
    lFileId = CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    strPrintName = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
    Select Case Val(vfgThis.TextMatrix(vfgThis.Row, mCol.�༭��ʽ))
        Case 0
            Set mfrmPrintPreview = New frmPrintPreview
            r = zlCommFun.ShowMsgBox("����Ԥ��/��ӡ", "��ѡ����Ԥ��/��ӡ�ĸ�ʽ��", "!���ո�ʽ(&F),ԭʼ��ʽ(&O),ȡ��(&C)", Nothing)
            If r = "���ո�ʽ" Then
                blnOrigMode = False
            ElseIf r = "ԭʼ��ʽ" Then
                blnOrigMode = True
            Else
                Exit Sub
            End If
            mfrmPrintPreview.DoMultiDocPreview Me, cprסԺ����, mlngPatiId, mlngPageId, _
                        vfgThis.Cell(flexcpText, vfgThis.Row, mCol.��������), vfgThis.Cell(flexcpText, vfgThis.Row, mCol.ҳ����), _
                        lFileId, Not blnPreview, blnOrigMode, , mblnMoved, mlngAdviceID, , IIf(InStr(mstrPrivs, "ȡ����ӡ") > 0, 0, 1)    'û��"ȡ����ӡ"Ȩ�޲������ظ���ӡ�������������ӡ����
            Unload mfrmPrintPreview 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
            Set mfrmPrintPreview = Nothing
            If Not blnPreview Then RefreshList 'ֱ�Ӵ�ӡ�ڴ�ˢ��
        Case 1
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, False, 0, cprPF_סԺ, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "������ӡ") > 0
            mObjTabEprView.zlPrintDoc Me, blnPreview, strPrintName
        Case 2
'            ��Ⱦ���Ѷ���ҳ�棬��Ҫ����ʾ
    End Select
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", strPrintName
End Sub

Private Sub modelsApply()
    Dim frmModels As New frmEPRModelsMan, strPrivs As String
    If frmModels.Showfrm(Me, mlngPatiId, mlngPageId, mlngDeptId, gstrPrivsEpr) Then RefreshList
End Sub
Private Function EprIsCommit() As String
'��|�ָ���ʽ����,״̬Ϊ0 ������ 1 �����ֱ���� ����|ɾ��|����

Dim rsTemp As ADODB.Recordset, intNew As Integer, intDel As Integer, intMod As Integer
    gstrSQL = "Select ����״̬ From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)

    Select Case NVL(rsTemp!����״̬, 0)
        Case 0
            intNew = 1: intDel = 1: intMod = 1
        Case 1 '�ȴ����
            intNew = 0: intDel = 0: intMod = 0
        Case 2 '�ܾ����
            intNew = 1: intDel = 1: intMod = 1
        Case 3 '�������
            intNew = 0: intDel = 0: intMod = 0
        Case 4 '��鷴��
            intNew = 0: intDel = 0: intMod = 1
        Case 5 '���鵵
            intNew = 0: intDel = 0: intMod = 0
        Case 6 '�������
            intNew = 0: intDel = 0: intMod = 1
        Case 13 '���ڳ��
            intNew = 1: intDel = 1: intMod = 1
        Case 14 '��鷴��
            intNew = 1: intDel = 1: intMod = 1
        Case 16 '�������
            intNew = 1: intDel = 1: intMod = 1
        Case Else
            intNew = 0: intDel = 0: intMod = 0
    End Select
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
End Function
Private Function GetEprSign(ByVal lngFileID As Long)
'��ȡ������ʷǩ����¼
Dim rsTemp As ADODB.Recordset, strSign As String
    gstrSQL = "Select ��ʼ�� As �汾, Decode(Ҫ�ر�ʾ, 3, '����ҽʦ', 2, '����ҽʦ', '����ҽʦ') || '���' || Decode(��ʼ��, 1, 'ǩ��', '�޶�') As ����," & vbNewLine & _
                "       Decode(Nvl(Instr(�����ı�, ';'), 0), 0, �����ı�, Substr(�����ı�, 1, Instr(�����ı�, ';') - 1)) As ��Ա," & vbNewLine & _
                "       RTrim(Substr(��������, Instr(��������, ';', 1, 4) + 1)) As ʱ��" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� = 8 Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����¼", lngFileID)
    Do Until rsTemp.EOF
        strSign = strSign & "�� " & Rpad(NVL(rsTemp!��Ա), 8) & "�� " & Rpad(NVL(rsTemp!ʱ��), 19) & " ��" & NVL(rsTemp!����) & vbCrLf
        rsTemp.MoveNext
    Loop
    GetEprSign = strSign
End Function
Private Sub PrintCancel(ByVal lngRecordId As Long)
'ȡ����Ǵ�ӡ
On Error GoTo errHand
    gstrSQL = "Zl_���Ӳ�����ӡ_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    vfgThis.Cell(flexcpData, vfgThis.Row, mCol.��ǰ���) = ""
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.��ӡ) = ""
    Set vfgThis.Cell(flexcpPicture, vfgThis.Row, mCol.ҳ������) = Nothing
    Set vfgThis.Cell(flexcpPicture, vfgThis.Row, mCol.��������) = Nothing
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function EprPrinted(ByVal lngRecordId As Long, Optional strPrintInfo As String) As Boolean
'��鵱ǰ������¼�Ƿ��Ѿ���ӡ��
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    '��Ҫ�������Ӳ�����¼����ӡ�ˣ���ӡʱ�䣩��������ʷ���ݲ�ת�ƣ���¼�������ϲ�ѯ
    gstrSQL = "Select ��ӡ��, ��ӡʱ�� From ���Ӳ�����ӡ Where �ļ�id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select ��ӡ��, ��ӡʱ�� From ���Ӳ�����¼ Where ID = [1] And ��ӡ�� is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.EOF Then Exit Function
    
    Do Until rsTemp.EOF
        strPrintInfo = strPrintInfo & vbCrLf & "��ӡ�ˣ�" & Rpad(rsTemp!��ӡ��, 8) & "��ӡʱ�䣺" & Format(rsTemp!��ӡʱ��, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function EprWriteMSG() As Boolean
Dim rsTemp As New ADODB.Recordset, strMsg As String
On Error GoTo errHand
    gstrSQL = "Select �ļ�ID ID,������� || '-' || �������� ����, ����ʱ��, ����" & vbNewLine & _
                "From ���Ӳ���ʱ��" & vbNewLine & _
                "Where ����id = [1] And ��ҳid = [2] And ����id =[3] And ������Դ = 2 And (Nvl(��ɼ�¼id, 0) = 0 And ���ʱ�� Is Null)" & vbNewLine & _
                "Order By ����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId)
    
    Do Until rsTemp.EOF
        strMsg = strMsg & "����<" & Rpad(rsTemp!���� & ">", 31) & "��δ��д���������ʱ��:" & Format(rsTemp!����ʱ��, "yyyy-MM-dd hh:mm") & "  " & _
                        IIf(NVL(rsTemp!����, 0) = 0, "����", "����") & "�Ǳ�����д�ģ����飡" & vbCrLf
        rsTemp.MoveNext
    Loop
    
    '����̫�࣬�������ʾ���ܿ���,ֻ��ʾʮ��
    If UBound(Split(strMsg, vbCrLf)) > 9 Then
        strMsg = Mid(strMsg, 1, InStr(710, strMsg, vbCrLf))
        strMsg = strMsg & String(32, Asc("-")) & "���»��ж�����¼��"
    End If
    
    If MsgBoxD(Me, strMsg & vbCrLf & "ѡ<��>������ѡ<��>ȡ����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        EprWriteMSG = False
    Else
        EprWriteMSG = True
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function TimeLimitOut() As Boolean
'����:����Ƿ���ת�ƣ���Ժ��Ԥ��Ժ�������������¼��Ͳ�¼ʱ��
Dim rsTemp As New ADODB.Recordset, lngTimeLimit As Long, strReturn As String
    If mintState = 3 Or mintState = 4 Then Exit Function
    
    gstrSQL = "Select Decode(��ֹԭ��, 1, '��Ժ', 3, 'ת��', 10, 'Ԥ��Ժ') �¼�, ��ֹʱ��,Trunc((Sysdate - ��ֹʱ��) * 24, 5) ��ǰʱ��" & vbNewLine & _
                "From ���˱䶯��¼" & vbNewLine & _
                "Where ID = (Select Nvl(Max(ID), 0)" & vbNewLine & _
                "            From ���˱䶯��¼" & vbNewLine & _
                "            Where ����id = [1] And ��ҳid = [2] And ��ֹʱ�� Is Not Null And ��ֹԭ�� In (1, 3, 10))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����䶯��¼", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    
    lngTimeLimit = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", 100))
    
    If rsTemp!��ǰʱ�� > lngTimeLimit Then
        If rsTemp!�¼� = "ת��" Then
            strReturn = rsTemp!�¼� & "|" & lngTimeLimit
            gstrSQL = "Select ��Ժ����id From ������ҳ Where ����id = [1] And ��ҳid = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", mlngPatiId, mlngPageId)
            If mlngDeptId = rsTemp!��Ժ����ID Then strReturn = "" 'ת�ƺ���ת�����������������ʱ������
        Else
            strReturn = rsTemp!�¼� & "|" & lngTimeLimit
        End If
    End If
    
    If strReturn <> "" Then
        MsgBox "�ò����Ѿ�" & Split(strReturn, "|")(0) & ",���ҳ����趨��" & Split(strReturn, "|")(1) & "Сʱ��¼ʱ��,������䶯������", vbInformation, gstrSysName
        TimeLimitOut = True
    End If
End Function
Private Function ExportAll() As Boolean
'���ܣ������ò�������ȫ��ʽ����ΪRTF
'���裺1 ָ��Ŀ¼
'     2 ���ļ����������������Ϊһ���ļ������뵽�ؼ�
'     3 ˢ�����ݶ���
'     4 ȥ���ؼ���
'     5 ����ΪRTF������Ϊ����(סԺ��)_��������
Dim strFile As String, strName As String, strPath As String, j As Long
Dim rsTemp As New ADODB.Recordset, strPage As String, lngLen As Long, blnExport As Boolean

    On Error GoTo errHand

    'ָ��Ŀ¼
    strPath = zl9ComLib.OS.OpenDir(Me.hwnd, "ָ������Ŀ¼")
    If strPath = "" Then
        MsgBox "ȡ��ָ������Ŀ¼������ʧ�ܣ�", vbExclamation, gstrSysName
        ExportAll = False: Exit Function
    End If
    Call zlCommFun.ShowFlash("���Եȣ����ڵ����ļ�", Me)
    
    gstrSQL = "Select a.סԺ��, a.���� From ������Ϣ a Where a.����id = [1]" 'ָ�������ļ�ǰ�
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId)
    strName = rsTemp!���� & "(סԺ��_" & rsTemp!סԺ�� & ")"
    
    strPath = gobjFSO.BuildPath(strPath, rsTemp!����) 'ָ��Ŀ¼�µ���Ŀ¼
    If Not gobjFSO.FolderExists(strPath) Then gobjFSO.CreateFolder strPath '������������Ŀ¼

    
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtBuff.ForceEdit = True
    gfrmPublic.edtPublic.Freeze
    gfrmPublic.edtBuff.Freeze
    For j = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(j, mCol.�༭��ʽ) = 0 Then
            '��ȡRTF��ˢ�����ݶ���
            If vfgThis.RowOutlineLevel(j) = 1 Then '�����ǰ������һ�е�ҳ��������ͬ����׷�ӣ����򵥶���
                Call ReadRTF(gfrmPublic.edtBuff, Val(vfgThis.TextMatrix(j, mCol.ID)), True, mblnMoved)
                gfrmPublic.edtBuff.SelectAll
                gfrmPublic.edtBuff.CopyWithFormat
                lngLen = Len(gfrmPublic.edtBuff.Text)
                If gfrmPublic.edtPublic.Range(lngLen - 2, lngLen).Text = vbCrLf Then '��β������
                    gfrmPublic.edtPublic.Range(lngLen - 2, lngLen).Font.Hidden = False
                Else
                    gfrmPublic.edtPublic.Range(lngLen, lngLen).Text = vbCrLf
                    gfrmPublic.edtPublic.Range(lngLen, lngLen + 2).Font.Hidden = False
                End If
                gfrmPublic.edtPublic.PasteWithFormat
            Else
                strPage = vfgThis.TextMatrix(j, mCol.ҳ������)
                Call ReadRTF(gfrmPublic.edtPublic, Val(vfgThis.TextMatrix(j, mCol.ID)), True, mblnMoved)
            End If
            
            
            blnExport = False
            If j = vfgThis.Rows - 1 Then
                blnExport = True
            ElseIf vfgThis.RowOutlineLevel(j + 1) = 0 Then
                blnExport = True
            End If
            
            If blnExport Then
                '������йؼ���
                Dim i As Long
                Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
                i = 0
                bFinded = FindNextAnyKey(gfrmPublic.edtPublic, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Do While bFinded
                    gfrmPublic.edtPublic.Range(lKSS, lKSE) = ""
                    gfrmPublic.edtPublic.Range(lKSS + lKES - lKSE, lKSS + lKES - lKSE + 16) = ""
                    i = lKSS + (lKES - lKSE)
                    bFinded = FindNextAnyKey(gfrmPublic.edtPublic, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Loop
                
                gfrmPublic.edtPublic.SaveDoc (strPath & "\" & strName & "_" & strPage & Format(vfgThis.TextMatrix(j, mCol.����ʱ��), "yyyymmddHHmmss") & ".rtf")
            End If
        End If
    Next
    gfrmPublic.edtPublic.ForceEdit = False
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtPublic.UnFreeze
    gfrmPublic.edtBuff.UnFreeze
    Unload gfrmPublic
    Call zlCommFun.StopFlash
    MsgBox "�ɹ������ļ���Ŀ¼ [" & strPath & "]��!", vbInformation, gstrSysName
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub EprArchive()
Dim strState As String, rsTemp As New ADODB.Recordset, strInfo As String

    On Error GoTo errHand
    gstrSQL = "Select Decode(��Ժ����, Null, Decode(״̬, 3, 'Ԥ��Ժ', '��Ժ'), '��Ժ') As ����״̬" & vbNewLine & _
                "From ������ҳ" & vbNewLine & _
                "Where ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�ǰ״̬", mlngPatiId, mlngPageId)
    strState = rsTemp!����״̬
    
    With vfgThis
        If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
            If Not EprWriteMSG Then Exit Sub
            If strState = "��Ժ" Then
                strInfo = "��Ľ��÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���鵵��"
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0)"
            Else
                strInfo = "�����Ѿ�" & strState & "��Ҫ�����˱���סԺȫ��סԺ�����鵵��" & vbCrLf _
                        & "  ѡ���ǡ����鵵���˱���ȫ��������" & vbCrLf _
                        & "  ѡ�񡰷񡱣����鵵�÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "����"
                Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                Case vbYes: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0,1)"
                Case vbNo: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0)"
                Case Else: Exit Sub
                End Select
            End If
        Else
    
            If Split(EprIsCommit, "|")(2) = 0 Then
                MsgBox "�ò��˲������ύ��飬���ܳ�������ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strInfo = "��Ҫ�����ò��˱���סԺ�����ѹ鵵סԺ������" & vbCrLf _
                    & "  ѡ���ǡ��������ò��˱���סԺ�����ѹ鵵סԺ������" & vbCrLf _
                    & "  ѡ�񡰷񡱣��������÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���Ĺ鵵��"
            Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
            Case vbYes: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",1,1)"
            Case vbNo: gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",1)"
            Case Else: Exit Sub
            End Select
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function GetPhysicians() As String
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If mlngPatiId = 0 Then Exit Function
    
    gstrSQL = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ" & vbNewLine & _
            "From ���˱䶯��¼" & vbNewLine & _
            "Where ����id = [1] And ��ҳid = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽʦ", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    GetPhysicians = ";" & NVL(rsTemp!����ҽʦ) & ";" & NVL(rsTemp!����ҽʦ) & ";" & NVL(rsTemp!����ҽʦ) & ";"
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function UserNewEMR() As Boolean
Dim rsTemp As New ADODB.Recordset, lngDeptId As Long
    On Error GoTo errHand
    gstrSQL = "Select ID From ���Ӳ�����¼ Where ����ID=[1] and ��ҳID=[2] and ��������=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���д��", mlngPatiId, mlngPageId)
    If Not rsTemp.EOF Then Exit Function 'д���ϲ���
    
    gstrSQL = "Select ��Ժ����ID From ������ҳ Where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���˵�ǰ����", mlngPatiId, mlngPageId)
    lngDeptId = NVL(rsTemp!��Ժ����ID, 0)
    If lngDeptId = 0 Then lngDeptId = mlngDeptId
    
    On Error Resume Next
    gstrSQL = "Select ����ID From �°没�����ÿ��� Where ����ID=[1]" 'û������������ù�����ѯ
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鱾���Ƿ�����", lngDeptId)
    If Err.Number <> 0 Then Err.Clear: Exit Function  'û�п��Ʊ�
    
    If rsTemp.EOF Then Exit Function '�б�����û����
    
    UserNewEMR = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function DisplayContent(ByVal lngId As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHand
    strSQL = "Select  �Ǽ���, �Ǽ�ʱ��,��������,������, ����ʱ��,�������˵��  From �������淴�� where �ļ�ID = [1] order by �Ǽ�ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
    If rsTemp.RecordCount = 0 Then
        DisplayContent = False
        Exit Function
    End If

    With vsfFeedback
        .Clear
        .Cols = 6
        .Rows = 1
        .ColWidth(0) = .Width / 10
        .ColWidth(1) = .Width / 10 * 2
        .ColWidth(2) = .Width / 10 * 3
        .ColWidth(3) = .Width / 10
        .ColWidth(4) = .Width / 10
        .ColWidth(5) = .Width / 10 * 2
        .TextMatrix(0, 0) = "�Ǽ���"
        .TextMatrix(0, 1) = "�Ǽ�ʱ��"
        .TextMatrix(0, 2) = "��������"
        .TextMatrix(0, 3) = "������"
        .TextMatrix(0, 4) = "����ʱ��"
        .TextMatrix(0, 5) = "�������˵��"
    End With
    
    Do Until rsTemp.EOF
        With vsfFeedback
            .Rows = .Rows + 1
            .ROWHEIGHT(.Rows - 1) = 350
            .TextMatrix(.Rows - 1, 0) = NVL(rsTemp!�Ǽ���)
            If IsDate(rsTemp!�Ǽ�ʱ�� & "") Then
                .TextMatrix(.Rows - 1, 1) = Format(rsTemp!�Ǽ�ʱ��, "yy/mm/dd HH:mm")
            Else
                .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!�Ǽ�ʱ��)
            End If
            .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!�Ǽ�ʱ��)
            .TextMatrix(.Rows - 1, 2) = NVL(rsTemp!��������)
            .TextMatrix(.Rows - 1, 3) = NVL(rsTemp!������)
            If IsDate(rsTemp!�Ǽ�ʱ�� & "") Then
                .TextMatrix(.Rows - 1, 1) = Format(rsTemp!����ʱ��, "yy/mm/dd HH:mm")
            Else
                .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!����ʱ��)
            End If
            .TextMatrix(.Rows - 1, 5) = NVL(rsTemp!�������˵��)
        End With
        rsTemp.MoveNext
    Loop
    DisplayContent = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfFeedback_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '�رշ�������鿴��
        If vsfFeedback.Visible Then
            vsfFeedback.Visible = False
        End If
    End If
End Sub

Public Function GetFormOperation() As String
'��¼����ѡ����Ϣ����Ϊ����վ���л�ҳ��ʱ���ͷ��˶��󣬻�����ʱ���³�ʼ��ˢ�µġ�
    GetFormOperation = mlngCurId
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'�ָ�����ѡ����Ϣ������վ��ˢ��֮ǰ����
    mlngCurId = Val(strValue)
End Sub

Private Sub vsfFeedback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    With vsfFeedback
        If .MouseRow >= 0 And .MouseCol >= 0 Then
            Call zlCommFun.ShowTipInfo(.hWnd, .TextMatrix(.MouseRow, .MouseCol), True, True)
        End If
    End With
End Sub
