VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockOutEPRs 
   BorderStyle     =   0  'None
   Caption         =   "���ﲡ����¼"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsColumn 
      Height          =   3480
      Left            =   1215
      TabIndex        =   1
      Top             =   1875
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
      FormatString    =   $"frmDockOutEPRs.frx":0000
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
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   315
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frmDockOutEPRs.frx":004E
         ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2655
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2745
      Left            =   225
      TabIndex        =   0
      Top             =   675
      Width           =   7890
      _cx             =   13917
      _cy             =   4842
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
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
      Begin MSComctlLib.ImageList imgThis 
         Left            =   0
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":059C
               Key             =   "��д"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":0B36
               Key             =   "�޶�"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":10D0
               Key             =   "�鵵"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":166A
               Key             =   "ת��"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   720
      Top             =   4875
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockOutEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------
'���峣��
'-----------------------------------------------------
Private Enum mCol
    ��־ = 0: ID: ��������: ��������: ������: ����ʱ��: ������: ���ʱ��: ��ǰ�汾: ǩ������: ��ǰ���: �鵵��: �鵵����: ����ID: ������: ����״̬: ��ӡ��: ��ӡʱ��: �༭��ʽ: �걨״̬
End Enum

Const conPane_Content = 1
Const conPane_New = 2
Private mstrColWidthConfig As String

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '�̳��ĵ�����ġ������ϲο��¼���
Public Event RequestRefresh() 'Ҫ��������ˢ��
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ���߶Ա�����(1250)��Ȩ�޴�
Private mblnSearch As Boolean   '��ǰʹ�����Ƿ�߱���������(1273)Ȩ
Private mlngPatiId As Long      '����id
Private mlngPageId As Long      '��ҳid
Private mlngDeptId As Long      '��ǰ��������id
Private mblnEdit As Boolean     '�Ƿ��������
Private mblnMoved As Boolean    '�Ƿ�ת��
Private mlngAdviceID As Long    'ҽ��ID
Private mblnOutDoc As Boolean   '�Ƿ����������ݲ���

Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Attribute mfrmMonitor.VB_VarHelpID = -1
Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR            '���ʽ�����༭��
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Private mbln��Ⱦ�� As Boolean              '��Ⱦ�����濨�ڲ�����ɽ���֮��Ҳ�ǿ����޸ĵ�

Private mcbsThis As Object          'CommandBar�ؼ�
Private mlngVersion As Long         'ѡ�е��ļ��汾��
Private mblnDisease As Boolean      '�Ƿ�ӵ����1249ģ���Ȩ��

Private Sub InitColumnSelect()
    On Error Resume Next
    '���ܣ�����ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vfgThis
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.��������, mCol.������, mCol.����ʱ��, mCol.������, mCol.���ʱ��, mCol.��ǰ���, mCol.������
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

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane.ID
    Case conPane_Content
        Cancel = True
    Case conPane_New
        Select Case Action
        Case PaneActionClosing, PaneActionClosed: Cancel = False
        Case Else: Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
        Item.Handle = mfrmContent.hwnd
    Case conPane_New
        If mfrmNew Is Nothing Then Set mfrmNew = New frmDockEPRNew
        Item.Handle = mfrmNew.hwnd
    End Select
End Sub

Private Sub dkpMan_Resize()
    Dim lngScaleLeft As Long, lngScaleTop  As Long, lngScaleRight  As Long, lngScaleBottom  As Long
    Call Me.dkpMan.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Err = 0: On Error Resume Next
    With Me.vfgThis
        .Left = lngScaleLeft: .Width = lngScaleRight - lngScaleLeft
        .Top = lngScaleTop: .Height = lngScaleBottom - .Top
        .ZOrder 0
    End With
    fraColSel.Move Me.vfgThis.Left + 50, Me.vfgThis.Top + 50
    fraColSel.ZOrder 0
    vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
    vsColumn.ZOrder 0
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

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
Dim i As Integer
    For i = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(i, mCol.ID) = lngRecordId Then
            vfgThis.Cell(flexcpText, i, mCol.��ӡ��) = gstrUserName
            vfgThis.Cell(flexcpText, i, mCol.��ӡʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm")
            Exit For
        End If
    Next
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    If mblnOutDoc Then RaiseEvent RequestRefresh
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

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split("270;0;0;2200;800;1600;800;1600;0;0;3000;0;0;0;1200;0;800;1600;0", ";")
            vfgThis.ColWidth(lngCol) = T(lngCol)
            vfgThis.ColHidden(lngCol) = False
        Else
            vfgThis.ColWidth(lngCol) = 0
            vfgThis.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To 18
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

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
 
Private Sub Form_Load()
    Dim intType As Integer, lngFontSize As Long
    
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1250)
    
    mblnOutDoc = Val(zlDatabase.GetPara("��ʾ�����������", glngSys, 1260, 0, , , intType)) = 1
    
    mstrColWidthConfig = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
        "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0")
    lngFontSize = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", 9)
    vfgThis.FontSize = lngFontSize
    Dim panContent As Pane, panNew As Pane
    mlngPatiId = -1: mlngPageId = -1
    
    Set mfrmContent = New frmDockEPRContent
    Set panContent = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, Nothing)
    panContent.Title = "��������"
    panContent.Options = PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmNew = New frmDockEPRNew
    Set panNew = dkpMan.CreatePane(conPane_New, 200, 400, DockRightOf, Nothing)
    panNew.Title = "��������"
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    panNew.Close
    
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Me.dkpMan.Options.ThemedFloatingFrames = True
    mlngVersion = 1  'Ĭ��Ϊ��1��

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String, i As Long
    If vfgThis.Cols = 19 Then
        For i = 0 To 18
            strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
        Next
    Else
        strCols = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
            "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0")
    End If
    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", vfgThis.FontSize
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mobjDoc = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim frmThis As Form, bFinded As Boolean

        
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, 1250, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
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
    strSQL = "Select ���� From �����ļ��б� Where  ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    If rs!���� < 0 Then
        '���ⲡ������������
        Exit Sub
    ElseIf rs!���� = 2 Then '���ʽ�༭��
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_����, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_����, cprET_�������༭, FileId, True, 0, cprPF_����, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, , InStr(gstrPrivsEpr, "������ӡ") > 0, Val(gstrESign)
        End If
    ElseIf rs!���� = 4 Then '��Ⱦ�����濨�༭��
'        �Ѷ���ҳ��
    Else
        For Each frmThis In Forms
            If TypeName(frmThis) = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.����ID = mlngPatiId _
                        And .EPRPatiRecInfo.������Դ = cprPF_���� And .EPRPatiRecInfo.��ҳID = mlngPageId _
                        And .EPRPatiRecInfo.����ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                    End If
                End With
            End If
        Next
        If bFinded = False Then
            Set mobjDoc = New cEPRDocument
            mobjDoc.InitEPRDoc cprEM_����, cprET_�������༭, FileId, cprPF_����, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
            mobjDoc.ShowEPREditor Me
            Me.dkpMan.Panes(conPane_New).Close
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjDoc_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cbrControl As CommandBarControl
    vfgThis.Row = IIf(vfgThis.MouseRow = -1, vfgThis.Rows - 1, vfgThis.MouseRow)
    If Button = vbRightButton And Not mcbsThis Is Nothing Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"):  cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngRecordId As Long, blnRTFFile As Boolean, byteEdit As Byte
    
    Me.dkpMan.Panes(conPane_New).Close
    Err = 0: On Error Resume Next
    With Me.vfgThis
        If .Cols < mCol.ID + 1 Then Exit Sub
        lngRecordId = Val(.TextMatrix(.Row, mCol.ID))
        byteEdit = Val(.TextMatrix(.Row, mCol.�༭��ʽ))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> Me.vfgThis.Row) Then
        Call mfrmContent.zlRefresh(lngRecordId, IIf(mblnEdit = False, "", mstrPrivs), , mblnMoved, blnRTFFile, byteEdit, True)
        If blnRTFFile Then
            If dkpMan.Panes(conPane_Content).Closed = True Then Call dkpMan.Panes(conPane_Content).Select
        ElseIf dkpMan.Panes(conPane_Content).Selected = True Then
            dkpMan.Panes(conPane_Content).Close
        End If
        Me.vfgThis.Tag = Me.vfgThis.Row
    End If
End Sub

'------------------------------------------------------------
'����Ϊ��������
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ"): cbrControl.STYLE = xtpButtonIconAndCaption
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "�����������(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "���˲�������(&S)")
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
        'Set cbrControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ", cbrControl.Index + 1)
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With

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
    End With
    
    '-----------------------------------------------------
    '��û����д����ʱ������Ȩ��״̬����ʾ���Ӵ���
    '-----------------------------------------------------
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.FixedRows, mCol.ID)) = 0 Then
        If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0) Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String, lFileId As Long, blnCanPrint As Boolean
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_NewItem Or Control.ID = conMenu_Edit_Archive Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML) Then '��ת������,�޸�,ɾ��,����,�鵵,��,�������������
        MsgBox "�ò��˵ı��ξ��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lFileId = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    bEditor = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.�༭��ʽ))
    blnCanPrint = IIf(Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.���ʱ��)) = "", InStr(1, gstrPrivsEpr, "δǩ����ӡ") > 0, InStr(1, gstrPrivsEpr, "������ӡ") > 0) And (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
    Select Case Control.ID
    Case conMenu_File_Open
        '�����Ķ�
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView
            If EprPrinted(lFileId) And InStr(mstrPrivs, "ȡ����ӡ") = 0 Then blnCanPrint = False ''�Ѿ���ӡ����û��ȡ����ӡȨ��,�������ظ���ӡ
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint, , mlngAdviceID
        ElseIf bEditor = 1 Then
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_����, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_����, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, blnCanPrint, Val(gstrESign)
            End If
        ElseIf bEditor = 2 Then
'            ��Ⱦ���Ѷ���ҳ��
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
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
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_ExportToXML
        '������XML�ļ�
        Dim strF As String
        dlgThis.Filename = "����_" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.��������) & _
            "(" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.ID) & "," & mlngVersion & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        On Error GoTo errHand
        strF = dlgThis.Filename
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If bEditor = 1 Then
            '���ʽ����
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, False, 0, cprPF_����, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument '��ͨסԺ����
            DocXML.InitAndOpenEPR lFileId, mlngVersion, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    Case conMenu_Edit_NewItem
        Me.dkpMan.Panes(conPane_New).Select
        Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
    Case conMenu_Edit_Modify
        If EprPrinted(lFileId) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
        If bEditor = 1 Then
            '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_����, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, True, 0, cprPF_����, _
                    mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "������ӡ") > 0, Val(gstrESign)
            End If
        ElseIf bEditor = 0 Then
            For Each frmThis In Forms
                If frmThis.Name = "frmMain" Then
                    With frmThis.Document
                        On Error Resume Next
                        If .EPRPatiRecInfo.ID = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 1) And .EPRPatiRecInfo.����ID = mlngPatiId _
                            And .EPRPatiRecInfo.������Դ = cprPF_���� And .EPRPatiRecInfo.��ҳID = mlngPageId _
                            And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            bFinded = True
                        End If
                    End With
                End If
            Next
            If bFinded = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, lFileId, cprPF_����, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
                mobjDoc.ShowEPREditor Me
            End If
        ElseIf bEditor = 2 Then
'            ��Ⱦ���Ѷ���ҳ��
        End If
    Case conMenu_Edit_Delete
        With Me.vfgThis
            strInfo = "���ɾ����ݡ�" & .TextMatrix(.Row, mCol.��������) & "����"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "��ǰ�����Ѵ�ӡ���������������ȷ���ٴβ�����ȡ����ӡ���ٽ��У�", vbInformation, gstrSysName: Exit Sub
            gstrSQL = "Zl_���Ӳ�����¼_Delete(" & .TextMatrix(.Row, mCol.ID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
            
            RaiseEvent RequestRefresh
        End With
    Case conMenu_Edit_Archive
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
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End With
    Case conMenu_Edit_NoPrint 'ȡ����ӡ���
        Call PrintCancel(lFileId)
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 1, mlngDeptId, 1, 1)
    Case conMenu_Tool_Search
        Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_View_Refresh:  Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '���ʽ������28δ��������ǩ�����
            'call
        End If
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnTmp As Boolean
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    With Me.vfgThis
        Select Case Control.ID
        Case conMenu_File_Open, conMenu_File_Excel, conMenu_File_RowPrint
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_NoPrint
            Control.Enabled = InStr(mstrPrivs, "ȡ����ӡ") > 0 And (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            If Control.Enabled Then Control.Enabled = Trim(.TextMatrix(.Row, mCol.��ӡ��)) <> ""
            If Control.Enabled Then Control.Enabled = mblnEdit
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "������ӡ") > 0)
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Or InStr(1, mstrPrivs, "�鵵�������") > 0)
            If Control.Enabled And (Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_ExportToXML) Then
                Control.Enabled = Val(.TextMatrix(.Row, mCol.�༭��ʽ)) <> 2
            End If
        Case conMenu_Edit_NewItem
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
        Case conMenu_Edit_Modify
            If mblnDisease Then
                Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
            Else
                If Val(.TextMatrix(.Row, mCol.��������)) = 5 Then
                    Control.Enabled = (mbln��Ⱦ�� And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
                Else
                    Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0)
                End If
                If Control.Enabled Then
                    blnTmp = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0) '�Ѿ������������Ĳ������ܴ���
                    If Not blnTmp Then
                        If Val(.TextMatrix(.Row, mCol.�걨״̬)) = 4 Or Val(.TextMatrix(.Row, mCol.�걨״̬)) = 5 Then
                            blnTmp = True
                        End If
                    End If
                    Control.Enabled = blnTmp
                End If
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
        Case conMenu_Edit_Archive
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "�����鵵") > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.����״̬)) <= 0)  '�Ѿ������������Ĳ������ܴ���
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.ǩ������)) <> 0)         '��ǰ�汾�Ѿ�ǩ����ɲſ��Թ鵵
            If Trim(.TextMatrix(.Row, mCol.�鵵��)) = "" Then
                Control.Caption = "�鵵": Control.Checked = False
            Else
                Control.Caption = "����": Control.Checked = True
            End If
        Case conMenu_Tool_Monitor
            Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "�������") > 0)
        Case conMenu_Tool_Search: Control.Enabled = mblnSearch
        Case conMenu_Tool_SignVerify
            Control.Enabled = Val(.TextMatrix(.Row, mCol.ID)) <> 0 And Trim(.TextMatrix(.Row, mCol.���ʱ��)) <> ""
       End Select
    End With
End Sub

Public Sub RefreshList()
    Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    RaiseEvent RequestRefresh
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-С(ȱʡ)��1-��
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long) As Long
    Dim lngCol As Long, lngRow As Long
    Dim strKind As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��Ⱦ������ As String
    Dim rs��Ⱦ As ADODB.Recordset
    Dim str���� As String
    
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '��ȡ�Ƿ񱾲������õ���ǩ��,���ұ����ûȡ��ʱ��ȡ
        gstrESign = getPassESign(0, lngDeptId)
    End If
    mblnDisease = (GetPrivFunc(glngSys, 1249) <> "")   'true-�����˼�������ģ��;false-�����ü�������ģ��
    
    mlngPatiId = lngPatiID: mlngPageId = lngPageId: mlngAdviceID = lngAdviceID
    mlngDeptId = lngDeptId: mblnEdit = blnEdit: mblnMoved = blnMoved
    
    vsColumn.Visible = False
    Me.vfgThis.Tag = ""
    
    If mblnDisease Then
        str���� = " r.�������� In (1,6) "
    Else
        str���� = " (r.�������� In (1,6) or (r.��������=5 And r.�༭��ʽ<>2)) "
    End If

    gstrSQL = "Select r.Id, r.��������, r.��������, r.������ As ������, To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, r.������," & _
            "        To_Char(r.���ʱ��, 'yy-mm-dd hh24:mi') As ���ʱ��, r.���汾 As ��ǰ�汾, r.ǩ������," & _
            "        Decode(r.���汾, 1, '', '�޶���') || r.������ || '��' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') ||" & _
            "        Decode(Nvl(r.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') As ��ǰ���, r.�鵵��, r.�鵵����, r.����id," & _
            "        d.���� As ������, r.����״̬,r.��ӡ��,To_Char(r.��ӡʱ��, 'yyyy-mm-dd hh24:mi') As ��ӡʱ��,Decode(r.�༭��ʽ,2,Decode(r.��������,1,0,r.�༭��ʽ),r.�༭��ʽ) as �༭��ʽ,null as �걨״̬" & _
            " From ���Ӳ�����¼ r, ���ű� d" & _
            " Where r.����id = d.Id And r.������Դ = 1 And " & str���� & " And r.����id = [1] And Nvl(r.��ҳid, 0) = [2]" & _
            " Order By r.��������, r.���, r.����ʱ��"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    If Not mblnDisease Then
        gstrSQL = "Select a.����״̬,b.id,C.ִ�в���ID,C.ִ��״̬ From �����걨��¼ a,���Ӳ�����¼ b, ���˹Һż�¼ c  where a.�ļ�id=b.id and b.��������=5" & vbNewLine & _
            "and b.����id=c.����id and b.��ҳid=c.id and  c.id=[1] and a.����״̬ in (4,5)"
        Set rs��Ⱦ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPageId)
        
        mbln��Ⱦ�� = False
        If mblnEdit Then
            mbln��Ⱦ�� = True
        ElseIf rs��Ⱦ.RecordCount > 0 Then
            If Val(rs��Ⱦ!ִ�в���ID) = lngDeptId And (Val(rs��Ⱦ!ִ��״̬) = 1 Or Val(rs��Ⱦ!ִ��״̬) = 2) Then
                mbln��Ⱦ�� = True
            End If
        End If
        For lngRow = 1 To rs��Ⱦ.RecordCount
            str��Ⱦ������ = str��Ⱦ������ & "," & rs��Ⱦ!ID
            rs��Ⱦ.MoveNext
        Next
    End If
    
    With Me.vfgThis
        .Clear
        Set .DataSource = rsTemp
        
        Dim T As Variant, i As Long
        On Error Resume Next
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) < 18 Then
            mstrColWidthConfig = "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0;0"
        Else
            For i = 0 To 18
                .ColWidth(i) = T(i)
            Next
        End If
        
        If .FixedRows > 0 Then .ROWHEIGHT(.FixedRows - 1) = .RowHeightMin
        .MergeRow(0) = True
        For lngCol = .FixedCols To .Cols - 1
            .FixedAlignment(lngCol) = flexAlignCenterCenter
        Next
        strKind = ""
        For lngRow = .FixedRows To .Rows - 1
            If strKind <> .TextMatrix(lngRow, mCol.��������) Then
                '����������
                If strKind <> "" Then .CellBorderRange lngRow, 0, lngRow, .Cols - 1, RGB(0, 0, 255), 0, 1, 0, 0, 0, 0
                strKind = .TextMatrix(lngRow, mCol.��������)
            End If
            If Val(.TextMatrix(lngRow, mCol.����״̬)) > 0 Then
                Set .Cell(flexcpPicture, lngRow, mCol.��־) = imgThis.ListImages("ת��").Picture
            ElseIf Trim(.TextMatrix(lngRow, mCol.�鵵��)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.��־) = imgThis.ListImages("�鵵").Picture
            ElseIf Val(.TextMatrix(lngRow, mCol.��ǰ�汾)) <= 1 Then
                Set .Cell(flexcpPicture, lngRow, mCol.��־) = imgThis.ListImages("��д").Picture
            Else
                Set .Cell(flexcpPicture, lngRow, mCol.��־) = imgThis.ListImages("�޶�").Picture
            End If
            If .ROWHEIGHT(lngRow) < .RowHeightMin Then .ROWHEIGHT(lngRow) = .RowHeightMin
            If str��Ⱦ������ <> "" Then
                If InStr(str��Ⱦ������ & ",", "," & Val(.TextMatrix(lngRow, mCol.ID)) & ",") > 0 Then
                    rs��Ⱦ.Filter = "id=" & Val(.TextMatrix(lngRow, mCol.ID))
                    If Not rs��Ⱦ.EOF Then
                        .TextMatrix(lngRow, mCol.�걨״̬) = Val(rs��Ⱦ!����״̬ & "")
                    End If
                End If
            End If
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        vfgThis.Tag = -1: .Row = 0 '��ʹvfgthis��ѡ���κ��У�����ʾ�κ����ݣ�����ѡ��ĳ��ʱ��ˢ��
        If rsTemp.RecordCount = 1 Then
            .Row = 1
        End If
        Call vfgThis_RowColChange
    End With
    
    Call InitColumnSelect '��ѡ����
    
    '-----------------------------------------------------
    '��û����д����ʱ������Ȩ��״̬����ʾ���Ӵ���
    '-----------------------------------------------------
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.FixedRows, mCol.ID)) = 0 Then
        If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0) Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlOpenDefaultEPR(Optional ByVal bytKind As Byte = 1) As Boolean
    '******************************************************************************************************************
    '���ܣ��Զ�����һ��ȱʡ��������ﲡ��
    '������bytKind=1��ʾ���ﲡ��;=2��ʾ�Ǽ��ﲡ��;3=����
    '˵���������ǰ�������в���������Ҫ�Զ�����
    '******************************************************************************************************************
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "������д") > 0) Then
        With vfgThis
            If .Rows = 2 And Val(.TextMatrix(1, mCol.ID)) = 0 Then
                zlOpenDefaultEPR = mfrmNew.zlOpenDefaultEPR(bytKind)
            End If
        End With
    End If
End Function

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
    Dim strSubhead1 As String, strSubhead2 As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select r.�����, r.����, r.�Ա�, r.����, r.�Ǽ�ʱ��, r.No From ���˹Һż�¼ r Where r.Id =[1] and r.��¼����=1  and r.��¼״̬=1"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���˹Һż�¼", "H���˹Һż�¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPageId)
    If Not rsTemp.EOF Then
        strSubhead1 = "�����:" & rsTemp!����� & "  ����:" & rsTemp!���� & "  �Ա�:" & rsTemp!�Ա�
        strSubhead2 = "����:" & Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd") & "(No:" & rsTemp!NO & ")"
    Else
        strSubhead1 = "": strSubhead2 = ""
    End If
    
    Err = 0: On Error GoTo 0
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead1)
    Call objAppRow.Add(strSubhead2)
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
'##s
'## ������  blnPreview  :�Ƿ���Ԥ��ģʽ
'################################################################################################################
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long, strPrintName As String
    
    lFileId = CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    strPrintName = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
    Select Case Val(vfgThis.TextMatrix(vfgThis.Row, mCol.�༭��ʽ))
        Case 0
            Set mfrmPrintPreview = New frmPrintPreview
            mfrmPrintPreview.DoMultiDocPreview Me, cpr���ﲡ��, , , vfgThis.Cell(flexcpText, vfgThis.Row, mCol.��������) _
                            , , lFileId, Not blnPreview, , , mblnMoved, mlngAdviceID, strPrintName, IIf(InStr(mstrPrivs, "ȡ����ӡ") > 0, 0, 1) 'û��"ȡ����ӡ"Ȩ�޲������ظ���ӡ�������������ӡ����
            Unload mfrmPrintPreview 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
            Set mfrmPrintPreview = Nothing
        Case 1
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, lFileId, False, 0, cprPF_����, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "������ӡ") > 0
            mObjTabEprView.zlPrintDoc Me, blnPreview, strPrintName
        Case 2
'            ��Ⱦ���Ѷ���ҳ��
    End Select
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", strPrintName
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
        strPrintInfo = strPrintInfo & vbCrLf & "��ӡ�ˣ�" & Rpad(rsTemp!��ӡ��, 5) & "��ӡʱ�䣺" & Format(rsTemp!��ӡʱ��, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
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
            intNew = 0: intDel = 0: intMod = 1
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
Private Sub PrintCancel(ByVal lngRecordId As Long)
'ȡ����Ǵ�ӡ
On Error GoTo errHand
    gstrSQL = "Zl_���Ӳ�����ӡ_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.��ӡ��) = ""
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.��ӡʱ��) = ""
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Function GetFormOperation() As String
'��¼����ѡ����Ϣ����Ϊ����վ���л�ҳ��ʱ���ͷ��˶��󣬻�����ʱ���³�ʼ��ˢ�µġ�
    GetFormOperation = ""
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'�ָ�����ѡ����Ϣ������վ��ˢ��֮ǰ����
End Sub
