VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockDisease 
   BorderStyle     =   0  'None
   Caption         =   "��Ⱦ�����濨"
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   4200
      ScaleHeight     =   3120
      ScaleWidth      =   8145
      TabIndex        =   1
      Top             =   120
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
            Picture         =   "frmDockDisease.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3480
         Left            =   735
         TabIndex        =   3
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
         FormatString    =   $"frmDockDisease.frx":054E
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
         Left            =   960
         TabIndex        =   4
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
            Picture         =   "frmDockDisease.frx":059C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   5
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
                  Picture         =   "frmDockDisease.frx":6DEE
                  Key             =   "��д"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockDisease.frx":7388
                  Key             =   "�޶�"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockDisease.frx":7922
                  Key             =   "�鵵"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockDisease.frx":7EBC
                  Key             =   "ת��"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockDisease.frx":8256
                  Key             =   "��ӡ"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   1335
      Left            =   1425
      TabIndex        =   0
      Top             =   3285
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   75
      Top             =   3045
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   480
      Top             =   4920
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDockDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#-----------------------------------------------------
'���峣��
'-----------------------------------------------------
Private Enum mCol
    ��־ = 0
    ���˿��� = 1
    ҳ������ = 2
    �������� = 3
    ������ = 4
    ����ʱ�� = 5
    ������ = 6
    ���ʱ�� = 7
    ��ǰ�汾 = 8
    ǩ������ = 9
    ��ǰ��� = 10
    �鵵�� = 11
    �鵵���� = 12
    ����ID = 13
    ����״̬ = 14
    ID = 15
    �༭��ʽ = 16
    ��ӡ = 17
    Ӥ�� = 18
    �걨״̬ = 19
    ������¼ = 20
    ���� = 21
End Enum

Const conDefColWidth = "270;0;1200;1600;800;1600;800;1600;0;0;3300;0;0;0;0;0;0;0;0;0;0;0"

Private Const conPane_List = 1
Private Const conPane_Content = 2
Private Const conPane_ReportCard = 3
Private mstrColWidthConfig As String
'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '�̳��ĵ�����ġ������ϲο��¼���
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs                   As String                   '��ǰʹ���߶Ա�����(1250)��Ȩ�޴�
Private mlngPatiId                  As Long                     '����id
Private mlngPageId                  As Long                     '��ҳid
Private mlngDeptId                  As Long                     '��ǰ��������id����һ���ǵ�ǰ���˿���
Private mbytFrom                    As Byte
Private mblnEdit                    As Boolean                  '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˿��Ҿ�����
Private mblnMoved                   As Boolean                  '�Ƿ������Ѿ�ת��
Private mintState                   As Integer                  '��clsDockInEPR
Private mstrPhysicians              As String                   '��������ҽʦ���ִ�
Private WithEvents mobjDoc          As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr                  As cTableEPR                '���ʽ�����༭��
Private mObjTabEprView              As cTableEPR
Private mcbsThis                    As Object                   'CommandBar�ؼ�
Private mlngVersion                 As Long                     'ѡ�е��ļ��汾��
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private WithEvents mfrmContent      As frmDockInContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mfrmMonitor                 As New frmDockEPRMonitor
Private mfrmTipInfo                 As New frmTipInfo
Private mobjInfection               As Object

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_List
            Item.Handle = picList.hWnd
        Case conPane_Content
            Item.Handle = mfrmContent.hWnd
        Case conPane_ReportCard
            Item.Handle = mobjInfection.zlGetForm.hWnd
    End Select
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
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
    mfrmTipInfo.ShowTipInfo picInfo.hWnd, strTipInfo, True
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
    Dim lngMouseRow As Long
    Dim lngMouseCol As Long
    Dim lngWidth As Long, lngHeight As Long, i As Long
    
    With vfgThis
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                If DisplayContent(Val(vfgThis.TextMatrix(lngMouseRow, mCol.ID))) Then
                    For i = 0 To mCol.������¼ - 1
                        lngWidth = lngWidth + vfgThis.ColWidth(i)
                    Next
                    For i = 0 To lngMouseRow
                        lngHeight = lngHeight + IIf(.ROWHEIGHT(i) < .RowHeightMin, .RowHeightMin, .ROWHEIGHT(i))
                    Next
                    With vsfFeedback
                        .Left = picList.Left + vfgThis.Left + lngWidth
                        .Top = picList.Top + vfgThis.Top + lngHeight
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
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    If Not mfrmPrintPreview Is Nothing Then Unload mfrmPrintPreview
    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
    Unload mobjInfection.zlGetForm
    Set mobjInfection = Nothing
    Set mfrmContent = Nothing
    Set mfrmMonitor = Nothing
    Set mobjDoc = Nothing
    Set mfrmPrintPreview = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mfrmTipInfo = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub Form_Load()
    Dim panList As Pane, panContent As Pane, panReportCard As Pane, lngFontSize As Long
    mlngPatiId = -1: mlngPageId = -1
    mstrColWidthConfig = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", conDefColWidth)
    lngFontSize = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", 9)
    vfgThis.FontSize = lngFontSize
    mstrPrivs = GetPrivFunc(glngSys, 1249)
    
    Set panList = dkpMan.CreatePane(conPane_List, 200, 50, DockTopOf, Nothing)
    panList.Title = "�����б�"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmContent = New frmDockInContent
    Set panContent = dkpMan.CreatePane(conPane_Content, 200, 300, DockBottomOf, Nothing)
    panContent.Title = "��������"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If

    Set panReportCard = dkpMan.CreatePane(conPane_ReportCard, 200, 300, DockBottomOf, Nothing)
    panReportCard.Title = "���濨����"
    panReportCard.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mObjTabEprView = New cTableEPR
    Call mObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDBUser)

    With dkpMan
        .SetCommandBars mcbsThis
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With

    mlngVersion = 1  'Ĭ��Ϊ��1��
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
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngRecordId As Long, byteEdit As Byte
    Dim ControlBar As Object
    Dim blnAllowDelete As Boolean
    On Error GoTo errHand

    With Me.vfgThis
        If .Rows <= 1 Then Exit Sub
        If .Cols < mCol.ID + 1 Then Exit Sub
        lngRecordId = Val(.TextMatrix(.Row, mCol.ID))
        byteEdit = Val(.TextMatrix(.Row, mCol.�༭��ʽ))
    End With
    If Not mcbsThis Is Nothing Then
        Set ControlBar = mcbsThis.FindControl(, conMenu_Edit_Delete, , True)
        zlUpdateCommandBars ControlBar
        If Not mcbsThis.FindControl(, conMenu_Edit_Delete, , True) Is Nothing Then
            blnAllowDelete = mcbsThis.FindControl(, conMenu_Edit_Delete, , True).Enabled
        End If
    End If
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> Me.vfgThis.Row) Then
        Me.Tag = "Refresh"                                              '����ˢ��̫�죬���򱨡��ܾ�Ȩ�ޡ�
        
        dkpMan.FindPane(conPane_Content).Close
        dkpMan.FindPane(conPane_ReportCard).Close
        
        If byteEdit = 2 Then
            dkpMan.ShowPane conPane_ReportCard
            mobjInfection.zlRefresh mlngPatiId, mlngPageId, lngRecordId, mblnMoved
        Else
            dkpMan.ShowPane conPane_Content
            Call mfrmContent.zlRefresh(lngRecordId, IIf(mblnEdit = False, "", mstrPrivs), mblnMoved, , byteEdit, blnAllowDelete)
        End If
        Me.Tag = ""
        Me.vfgThis.Tag = Me.vfgThis.Row
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zlDefCommandBars(ByVal cbsThis As Object)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Set mcbsThis = cbsThis
    Set mcbsThis.Icons = zlCommFun.GetPubIcons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "�����������(&M)")
    End With
    
    '����������
    '-----------------------------------------------------
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
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FSHIFT, VK_F5, conMenu_View_Refresh
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options
        .AddHiddenCommand conMenu_Edit_Archive
        .AddHiddenCommand conMenu_Edit_Untread
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String, lFileId As Long, blnCanPrint As Boolean
    Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_Audit Or Control.ID = conMenu_Edit_Archive Or _
                        Control.ID = conMenu_File_ExportToXML) Then '��ת������,�޸�,ɾ��,���,�鵵,�򿪲��������
        MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If

    lFileId = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    bEditor = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.�༭��ʽ))
    
    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            If EprPrinted(lFileId) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
                MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
                Exit Sub
            End If
            Call zlEPRPrint(True)
        Case conMenu_File_Print
            If EprPrinted(lFileId) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then '�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
                MsgBox "��ǰ�����Ѵ�ӡ���������ظ���ӡ��", vbInformation, gstrSysName
                Exit Sub
            End If
            Call zlEPRPrint(False)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_File_ExportAll
            Call ExportAll
        Case conMenu_File_ExportToXML
            '������XML�ļ�
            Dim strF As String
            dlgThis.Filename = "����_" & vfgThis.TextMatrix(vfgThis.Row, mCol.��������) & "(" & lFileId & "," & mlngVersion & ").xml"
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
                If bEditor = 1 Then
                    '���ʽ����
                    mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, False, 0, mbytFrom, mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved
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
            Call AddNewReport
        Case conMenu_Edit_Modify
            If mbytFrom = 2 And bEditor <> 2 Then               'סԺ
                If TimeLimitOut Then Exit Sub   '������¼ʱ�ޣ��������޸ģ����������
            ElseIf mbytFrom = 1 And bEditor <> 2 Then           '����
                 If EprPrinted(lFileId) Then
                    MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName
                    Exit Sub
                 End If
            End If
            
            '�������༭ģʽ
            With Me.vfgThis
                If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
                If bEditor = 1 Then
                    '���ʽ����
                    If Not mObjTabEpr Is Nothing Then
                        bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, mbytFrom, mlngDeptId)
                    End If
                    If bFinded = False Then
                        Set mObjTabEpr = New cTableEPR
                        mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, mbytFrom, _
                            mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, InStr(1, mstrPrivs, "������ӡ") > 0, Val(gstrESign)
                    End If
                ElseIf bEditor = 0 Then
                    'RichEPR����
                    For Each frmThis In Forms
                        If frmThis.Name = "frmMain" Then
                            On Error Resume Next
                            If frmThis.Document.EPRPatiRecInfo.ID = lFileId And frmThis.Document.EPRPatiRecInfo.����ID = mlngPatiId _
                                And frmThis.Document.EPRPatiRecInfo.������Դ = mbytFrom And frmThis.Document.EPRPatiRecInfo.��ҳID = mlngPageId _
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
                        mobjDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, lFileId, mbytFrom, mlngPatiId, CStr(mlngPageId), 0, mlngDeptId
                        mobjDoc.ShowEPREditor Me, InStr(1, mstrPrivs, "������ӡ") > 0
                    End If
                ElseIf bEditor = 2 Then
                    mobjInfection.OpenDoc Me, cprEM_�޸�, mlngPatiId, mlngPageId, mbytFrom, Val(vfgThis.TextMatrix(vfgThis.Row, mCol.Ӥ��)), mlngDeptId, lFileId
                End If
            End With
            zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
        Case conMenu_Edit_Delete
            If bEditor <> 2 Then
                If mbytFrom = 2 Then
                    If Split(EprIsCommit, "|")(1) = 0 Then
                        MsgBox "�ò��˲������ύ��飬����ɾ������ȡ���������ԣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                With Me.vfgThis
                    If EprPrinted(lFileId) Then
                        MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strInfo = "���ɾ����ݡ�" & .TextMatrix(.Row, mCol.��������) & "����"
                    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    gstrSQL = "Zl_���Ӳ�����¼_Delete(" & lFileId & ")"
                    On Error GoTo errHand
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    Err = 0: On Error GoTo 0
                    zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
                End With
            ElseIf bEditor = 2 Then
                If MsgBox("ȷ��Ҫɾ����ݱ��濨��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                gstrSQL = "Zl_���Ӳ�����¼_Delete(" & lFileId & ")"
                On Error GoTo errHand
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                Err = 0: On Error GoTo 0
                zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
            End If
        Case conMenu_Edit_Audit
            If TimeLimitOut Then Exit Sub '������¼ʱ�ޣ��������޸ģ����������
            If EprPrinted(lFileId) Then
                MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If bEditor = 1 Then
                '���ʽ����
                If Not mObjTabEpr Is Nothing Then
                    bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, mbytFrom, mlngDeptId)
                End If
                If bFinded = False Then
                    Set mObjTabEpr = New cTableEPR
                    mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_���������, lFileId, True, 0, mbytFrom, _
                        mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, , Val(gstrESign)
                End If
            ElseIf bEditor = 0 Then
                '���������ģʽ
                Dim frmAudit As Form, bFindedAudit As Boolean
                For Each frmAudit In Forms
                    If frmAudit.Name = "frmMain" Then
                        On Error Resume Next
                        If frmAudit.Document.EPRPatiRecInfo.ID = lFileId _
                            And frmAudit.Document.EPRPatiRecInfo.������Դ = mbytFrom And frmAudit.Document.EPRPatiRecInfo.����ID = mlngPatiId _
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
                    mobjDoc.InitEPRDoc cprEM_�޸�, cprET_���������, lFileId, mbytFrom, mlngPatiId, CStr(mlngPageId), , mlngDeptId
                    mobjDoc.ShowEPREditor Me, InStr(1, mstrPrivs, "������ӡ") > 0
                End If
            End If
        Case conMenu_Edit_Archive
            Call EprArchive
        Case conMenu_Edit_NoPrint 'ȡ����ӡ���
            If mbytFrom = 2 Then
                If Split(EprIsCommit, "|")(0) = 0 Then
                    MsgBox "�ò��˲������ύ��飬���ܳ�����ӡ����ȡ���������ԣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            Call PrintCancel(lFileId)
        Case conMenu_Tool_Monitor
            If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
            Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, 1, mintState)
        Case conMenu_View_Refresh
            zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
        Case conMenu_Help_Help
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Tool_SignVerify
            If bEditor = 0 Then
                Call VerifySignature(Me, lFileId, mblnMoved)
            End If
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
    Dim blnTmp As Boolean
    Dim lngFileID As Long
    Dim lngEditor As Long
    
    With Me.vfgThis
        lngFileID = Val(.TextMatrix(.Row, mCol.ID))
        lngEditor = Val(.TextMatrix(.Row, mCol.�༭��ʽ))
    
        Select Case Control.ID
            Case conMenu_File_Excel, conMenu_File_RowPrint
                Control.Visible = (lngEditor <> 2)
                If Control.Visible Then Control.Enabled = (lngFileID <> 0)
            Case conMenu_Edit_NoPrint
                'Control.Visible = (lngEditor <> 2)
                If Control.Visible Then Control.Enabled = InStr(mstrPrivs, "ȡ����ӡ") > 0 And (lngFileID <> 0)
                If Control.Enabled Then Control.Enabled = Trim(.TextMatrix(.Row, mCol.��ӡ)) <> ""
                If Control.Enabled Then Control.Enabled = mblnEdit
            Case conMenu_Edit_NewItem
                Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
                Control.Enabled = (lngFileID <> 0 And InStr(1, mstrPrivs, "������ӡ") > 0)
                If Control.Enabled Then Control.Enabled = IIf(Trim(.TextMatrix(.Row, mCol.���ʱ��)) = "", InStr(1, mstrPrivs, "δǩ����ӡ") > 0, True)
                If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
                If Control.Enabled Then Control.Enabled = (vfgThis.TextMatrix(vfgThis.Row, mCol.������) = gstrUserName Or InStr(1, mstrPrivs, "��������") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '������д���в�������Ȩ��,��������ҽʦ
                If Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_ExportToXML Then
                    If Control.Enabled Then Control.Enabled = lngEditor <> 2
                End If
            Case conMenu_File_ExportAll
                Control.Enabled = (lngFileID <> 0 And InStr(1, mstrPrivs, "������ӡ") > 0)
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "��������") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '������д���в�������Ȩ��,��������ҽʦ
            Case conMenu_Edit_Modify
                Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
                If Control.Enabled Then
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
                        If .TextMatrix(.Row, mCol.��������) = "�л����񹲺͹���Ⱦ�����濨" Then
                            Control.Enabled = (InStr(1, mstrPrivs, "���˲���") > 0 Or Trim(.TextMatrix(.Row, mCol.������)) = Trim(gstrUserName)) And InStr(",2,3", IIf(.TextMatrix(.Row, mCol.�걨״̬) = "", 0, .TextMatrix(.Row, mCol.�걨״̬))) = 0
                        End If
                    End If
                End If
            Case conMenu_Edit_Delete
                Control.Enabled = (lngFileID <> 0) And (mblnEdit And mlngPatiId > 0 And (InStr(1, mstrPrivs, "������д") > 0 Or InStr(1, mstrPrivs, "ǿ��ɾ��") > 0))
                If Control.Enabled And InStr(1, mstrPrivs, "ǿ��ɾ��") > 0 Then Exit Sub                        '�߱�ǿ��ɾ��Ȩ�ޣ��򲻽��к������ж�
                If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)          '�Ѿ������������Ĳ������ܴ���
                If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.����ID)))    '���Ʋ����ſ���ɾ
                If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.���ʱ��)) = "")         'δ��ɲ�������ɾ
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "���˲���") > 0 Or Trim(.TextMatrix(.Row, mCol.������)) = Trim(gstrUserName))
            Case conMenu_Edit_Audit
                Control.Visible = (mbytFrom = 2 And lngEditor <> 2)
                If Control.Visible Then Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "��������") > 0)
                If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)          '�Ѿ������������Ĳ������ܴ���
                If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.���ʱ��)) <> "")        '��ɲ����ſ�����
                If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.�鵵��)) = "")           'δ�鵵����������
            Case conMenu_Edit_Archive
                Control.Visible = (lngEditor <> 2)
                If Control.Visible Then Control.Enabled = (mblnEdit And mlngPatiId > 0)
                If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)          '�Ѿ������������Ĳ������ܴ���
                If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.ǩ������)) <> 0)          '��ǰ�汾�Ѿ�ǩ����ɲſ��Թ鵵
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����鵵") > 0)
                If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
                    Control.Caption = "�鵵": Control.Checked = False
                Else
                    Control.Caption = "����": Control.Checked = True
                End If
            Case conMenu_Tool_Monitor
                Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "�������") > 0)
            Case conMenu_Tool_SignVerify
                 Control.Visible = (lngEditor = 0)
                If Control.Visible Then Control.Enabled = lngFileID <> 0 And Trim(.TextMatrix(.Row, mCol.���ʱ��)) <> ""
        End Select
    End With
End Sub

Public Sub RefreshList()
    zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
End Sub

Private Sub InitColumnSelect()
'���ܣ�����ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    On Error Resume Next
    vsColumn.Rows = vsColumn.FixedRows
    With vfgThis
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.��������, mCol.������, mCol.����ʱ��, mCol.������, mCol.���ʱ��, mCol.��ǰ�汾, mCol.��ǰ���, mCol.Ӥ��
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
End Sub

Private Sub Initvfg()
    On Error Resume Next
    With vfgThis
        mfrmContent.Clear
        .Tag = ""
        .Clear
        .Rows = 1
        .Cols = 22
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
        .TextMatrix(0, mCol.����״̬) = "����״̬"
        .TextMatrix(0, mCol.ID) = "ID"
        .TextMatrix(0, mCol.�༭��ʽ) = "�༭��ʽ"
        .TextMatrix(0, mCol.��ӡ) = "��ӡ"
        .TextMatrix(0, mCol.�걨״̬) = "�걨״̬"
        .TextMatrix(0, mCol.Ӥ��) = "Ӥ��"
        .TextMatrix(0, mCol.������¼) = "������¼"
	.TextMatrix(0, mCol.����) = "����"
        .MergeCellsFixed = flexMergeFree
        .MergeCol(mCol.ҳ������) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        
        Dim T As Variant, i As Long '�����п�
        T = Split(conDefColWidth, ";")
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

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal bytFrom As Byte, ByVal lngDeptId As Long, ByVal blnMoved As Boolean, Optional ByVal blnEdit As Boolean = True, _
                            Optional ByVal intState As Integer) As Long
''������lngPageId סԺ����ҳID�����ﴫ�Һ�ID
    Dim lngCurId As Long            '��ǰ������¼ID
    Dim lngCurRow As Long           'ˢ�º�ѡ���кţ�Ĭ��Ϊ0����ѡ��
    Dim rsTemp As ADODB.Recordset, rsDis As ADODB.Recordset
    Dim lngCol As Long, lngRow As Long, i As Long
    Dim strKind As String
    Dim strReportIDs As String
    Dim str��Ⱦ������ As String
    Dim rs��Ⱦ As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim str���� As String
    Dim blnOneCard As Boolean
    
    lngCurId = IIf(mlngPatiId = lngPatiID, Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)), 0) '��ǰ����ˢ��ǰѡ����ID
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '��ȡ�Ƿ񱾲������õ���ǩ��,���ұ����ûȡ��ʱ��ȡ
        gstrESign = getPassESign(1, lngDeptId)
    End If
    
    mlngDeptId = lngDeptId
    mblnEdit = blnEdit
    mblnMoved = blnMoved
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mbytFrom = bytFrom
    mintState = intState
    
    vsColumn.Visible = False
    mstrPhysicians = GetPhysicians '��ȡ����ҽ������
    picInfo.Visible = False
    vsfFeedback.Visible = False
    
    Call Initvfg
    
    On Error GoTo errHand
    blnOneCard = Val(zlDatabase.GetPara("��Ⱦ�����濨һ��һ��", glngSys, 1277, 0)) = 1
    gstrSQL = " Select r.����id ���˿���, r.�������� As ҳ��, r.��������, r.������ As ������, To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, r.������," & vbNewLine & _
              "       To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.���汾 As ��ǰ�汾, r.ǩ������," & vbNewLine & _
              "       Decode(r.���汾, 1, '��д��', '�޶���') || r.������ || '��' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') ||" & vbNewLine & _
              "        Decode(Nvl(r.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') As ��ǰ���, r.�鵵��, r.�鵵����, r.����id, r.����״̬, r.Id, r.�༭��ʽ," & vbNewLine & _
              "       r.��ӡ�� As ��ӡ, r.Ӥ��" & vbNewLine & _
              " From ���Ӳ�����¼ R" & vbNewLine & _
              " Where r.������Դ = [3] And r.�������� = 5 And r.����id = [1] And r.��ҳid = [2] And r.�༭��ʽ In (0,1,2)" & vbNewLine & _
              " Order By r.�ļ�ID, r.����ʱ��"
              
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mbytFrom)
    
    gstrSQL = " Select a.����״̬, b.Id From �����걨��¼ A, ���Ӳ�����¼ B" & vbNewLine & _
              " Where a.�ļ�id = b.Id And b.�������� = 5 And b.����id  = [1] And b.��ҳid = [2] And a.����״̬ In (4, 5)"
    Set rs��Ⱦ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)

    For lngRow = 1 To rs��Ⱦ.RecordCount
        str��Ⱦ������ = str��Ⱦ������ & "," & rs��Ⱦ!ID
        rs��Ⱦ.MoveNext
    Next
    
    strKind = ""
    With vfgThis
        .ColWidth(mCol.�걨״̬) = 0
        .ColHidden(mCol.�걨״̬) = True
        .ColWidth(mCol.������¼) = 0
        .ColHidden(mCol.������¼) = True
	.ColWidth(mCol.����) = 0
        .ColHidden(mCol.����) = True
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
            .TextMatrix(rsTemp.AbsolutePosition, mCol.����״̬) = NVL(rsTemp!����״̬)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ID) = NVL(rsTemp!ID)
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
            If blnOneCard Then
                If rsTemp!�������� & "" = "�л����񹲺͹���Ⱦ�����濨" Then
                    gstrSQL = "Select a.Ҫ������||'-'||a.�����ı� as ���� From ���Ӳ������� a Where a.�ļ�id = [1] and a.������� between 20 and 30 and a.�����ı� is not null Order By a.������� desc"
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rsTemp!ID & ""))
                    If Not rs����.EOF Then str���� = rs����!���� & ""
                    .ColWidth(mCol.����) = 800
                    .ColHidden(mCol.����) = False
                    .TextMatrix(rsTemp.AbsolutePosition, mCol.����) = str����
                End If
            End If
            If strKind <> .TextMatrix(rsTemp.AbsolutePosition, mCol.��������) Then '����������
                If strKind <> "" Then Call .CellBorderRange(rsTemp.AbsolutePosition, 0, rsTemp.AbsolutePosition, .Cols - 1, RGB(125, 125, 125), 0, 1, 0, 0, 0, 0)
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
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'-------------------------------------------------
'����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
'����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'       strSubhead����ӡ�ĸ�����
'-------------------------------------------------
Private Sub zlRptPrint(ByVal bytMode As Byte)
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
            mfrmPrintPreview.DoMultiDocPreview Me, cprסԺ����, mlngPatiId, mlngPageId, 5, , _
                        lFileId, Not blnPreview, blnOrigMode, , mblnMoved, , , IIf(InStr(mstrPrivs, "ȡ����ӡ") > 0, 0, 1)    'û��"ȡ����ӡ"Ȩ�޲������ظ���ӡ�������������ӡ����
            Unload mfrmPrintPreview 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
            Set mfrmPrintPreview = Nothing
            If Not blnPreview Then RefreshList 'ֱ�Ӵ�ӡ�ڴ�ˢ��
        Case 1
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, False, 0, mbytFrom, mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, InStr(mstrPrivs, "������ӡ") > 0
            mObjTabEprView.zlPrintDoc Me, blnPreview, strPrintName
        Case 2
            strPrintName = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
            mobjInfection.PrintDoc Me, mlngPatiId, mlngPageId, lFileId, strPrintName
                        Call RefreshList
    End Select
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", strPrintName
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
    Dim rsTemp As ADODB.Recordset, strMsg As String
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
    strPath = zl9ComLib.OS.OpenDir(Me.hWnd, "ָ������Ŀ¼")
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
    Dim strInfo As String

    On Error GoTo errHand
       If mbytFrom = 1 Then
        With Me.vfgThis
            If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
                strInfo = "��Ľ��÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���鵵��"
            Else
                strInfo = "��ĳ����÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���Ĺ鵵��"
            End If
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & "," & IIf(Trim(.TextMatrix(.Row, mCol.�鵵��)) = "", 0, 1) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
        End With
    ElseIf mbytFrom = 2 Then
        With vfgThis
            If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
                If Not EprWriteMSG Then Exit Sub
                strInfo = "��Ľ��÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���鵵��"
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0)"
            Else
                If Split(EprIsCommit, "|")(2) = 0 Then
                    MsgBox "�ò��˲������ύ��飬���ܳ�������ȡ���������ԣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                strInfo = "��ĳ����÷ݡ�" & .TextMatrix(.Row, mCol.��������) & "���Ĺ鵵��"
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                gstrSQL = "Zl_���Ӳ�����¼_Archive(" & .TextMatrix(.Row, mCol.ID) & ",1)"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
        End With
    End If
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
    If KeyCode = vbKeyEscape Then       '�رշ�������鿴��
        If vsfFeedback.Visible Then
            vsfFeedback.Visible = False
        End If
    End If
End Sub

Private Sub AddNewReport()
    Dim rsTemp As ADODB.Recordset
    Dim lngFileID As Long
    Dim objDoc As cEPRDocument
    Dim bFinded As Boolean
    
    On Error GoTo errHand
    
    gstrSQL = " Select a.Id, a.����, a.���, a.����, a.����, a.˵��" & vbNewLine & _
              " From �����ļ��б� A" & vbNewLine & _
              " Where ���� = 5 And " & vbNewLine & _
              "      (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [1]))" & vbNewLine & _
              " Order By a.���"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ��ID", mlngDeptId)
    If rsTemp.RecordCount <= 0 Then
        Exit Sub
    Else
        If rsTemp.RecordCount = 1 Then
            lngFileID = Val(rsTemp!ID & "")
        ElseIf rsTemp.RecordCount > 1 Then
            If frmDiseaseFileList.ShowMe(Me, rsTemp, lngFileID, "��ѡ��Ҫ��ӵı��濨����") Then
                rsTemp.Filter = "ID=" & lngFileID
            Else
                Exit Sub
            End If
        End If
        If Val(rsTemp!���� & "") = 2 Then '���༭��
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lngFileID, mlngPatiId, mlngPageId, mbytFrom, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_����, cprET_�������༭, lngFileID, True, 0, mbytFrom, _
                mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, InStr(1, mstrPrivs, "������ӡ") > 0, Val(gstrESign)
            End If
        ElseIf Val(rsTemp!���� & "") = 4 Then '�̶���ʽ�ı��濨
            mobjInfection.OpenDoc Me, cprEM_����, mlngPatiId, mlngPageId, mbytFrom, Val(vfgThis.TextMatrix(vfgThis.Row, mCol.Ӥ��)), mlngDeptId, lngFileID, True
        Else
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_����, cprET_�������༭, lngFileID, mbytFrom, mlngPatiId, mlngPageId, 0, mlngDeptId, 0, False)
            Call objDoc.ShowEPREditor(Me, , vbModeless)
        End If
        zlRefresh mlngPatiId, mlngPageId, mbytFrom, mlngDeptId, mblnMoved, mblnEdit, mintState
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    With vsfFeedback
        If .MouseRow >= 0 And .MouseCol >= 0 Then
            Call zlCommFun.ShowTipInfo(.hWnd, .TextMatrix(.MouseRow, .MouseCol), True, True)
        End If
    End With
End Sub