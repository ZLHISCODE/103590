VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceInsure 
   BorderStyle     =   0  'None
   Caption         =   "֧��������"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vs֧����� 
      Height          =   1080
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2715
      _cx             =   4789
      _cy             =   1905
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceInsure.frx":0000
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
      ExplorerBar     =   2
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
      Begin VB.Image imgAdd 
         Height          =   240
         Left            =   -480
         Picture         =   "frmEInvoiceInsure.frx":00CC
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   360
      Top             =   2760
      Width           =   525
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "֧��������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceInsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs֧�����.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Public Sub RefreshData()
    Dim strSQL As String, i As Integer, j As Integer
    Dim rs֧����� As ADODB.Recordset
    Dim str�������� As String
    
    On Error GoTo errHandle
    strSQL = "Select c.���� As ��������, a.Id As ���մ���id, a.����, b.�������, b.��������" & vbNewLine & _
                "From ����֧������ A, ֧�������� B, ������� C" & vbNewLine & _
                "Where a.Id = b.���մ���id(+) And a.���� = c.���" & vbNewLine & _
                "Order By c.����, a.����"

    Set rs֧����� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    With vs֧�����
        .Clear 1
        .Rows = 2
        .OutlineBar = flexOutlineBarSymbolsLeaf
        .Subtotal flexSTClear
        .MultiTotals = True
        .SubtotalPosition = flexSTAbove
        .OutlineCol = .ColIndex("��������")
        .Rows = rs֧�����.RecordCount + 1
        j = 1
        For i = 1 To rs֧�����.RecordCount
            If rs֧�����!�������� <> str�������� Then
                str�������� = rs֧�����!��������
                .AddItem rs֧�����!��������, j
                .RowData(j) = 1
                .MergeCol(2) = True
                .RowOutlineLevel(j) = 1
                .IsSubtotal(j) = True
                j = j + 1
             End If
             .TextMatrix(j, .ColIndex("��������")) = rs֧�����!��������
            .TextMatrix(j, .ColIndex("���մ���id")) = Val(rs֧�����!���մ���id)
            .TextMatrix(j, .ColIndex("����")) = Nvl(rs֧�����!����)
            .TextMatrix(j, .ColIndex("�������")) = Nvl(rs֧�����!�������)
            .TextMatrix(j, .ColIndex("��������")) = Nvl(rs֧�����!��������)
            .RowOutlineLevel(j) = 2
             .IsSubtotal(j) = True
             rs֧�����.MoveNext
             j = j + 1
        Next
        .Outline 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHandler
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&E)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        cbrControl.BeginGroup = True
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&E)", cbrControl.index + 1)
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPopup()
    '��ʾ�����˵�
    Dim objPopup As CommandBarPopup
    Err = 0: On Error GoTo ErrHandler
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub AddNewInsure()
    '����֧��������
    Dim frmEdit As New frmEInvoiceInsureSet
    Dim blnRefresh As Boolean
    Dim str֧������ As String, str�������� As String
    Dim lng֧������ID As Long
    
    On Error GoTo errHandle
    With vs֧�����
        str�������� = .TextMatrix(.Row, .ColIndex("��������"))
        str֧������ = .TextMatrix(.Row, .ColIndex("����"))
        lng֧������ID = Val(.TextMatrix(.Row, .ColIndex("���մ���id")))
    End With
    If lng֧������ID = 0 Then Exit Sub
    Call frmEdit.ShowMe(Me, 0, str��������, lng֧������ID, str֧������, , , blnRefresh)
    If blnRefresh Then Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteInsure()
    'ɾ��֧��������
    On Error GoTo errHandle
    Dim lng����ID As Long, str������� As String
    Dim str�������� As String, strSQL As String
    With vs֧�����
        If .Row = 0 Then Exit Sub
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("���մ���id")))
        str�������� = .TextMatrix(.Row, .ColIndex("��������"))
        str������� = .TextMatrix(.Row, .ColIndex("�������"))
        
        If str������� = "" Then Exit Sub
        If MsgBox("��ȷ��Ҫɾ����������Ϊ��" & str�������� & "��,�������Ϊ��" & str������� & "����֧����������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Me.MousePointer = 11
        strSQL = "Zl_֧��������_Update(2," & lng����ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Me.MousePointer = 0
    End With

    Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub ModifyInsure()
    '�޸�֧��������
    Dim frmEdit As New frmEInvoiceInsureSet
    Dim blnRefresh As Boolean
    Dim str�������� As String, str֧������ As String
    Dim str������� As String, str�������� As String
    Dim lng֧������ID As Long
    
    On Error Resume Next
    With vs֧�����
        If .Row = 0 Then Exit Sub
            str�������� = .TextMatrix(.Row, .ColIndex("��������"))
            str֧������ = .TextMatrix(.Row, .ColIndex("����"))
            lng֧������ID = Val(.TextMatrix(.Row, .ColIndex("���մ���id")))
            str������� = .TextMatrix(.Row, .ColIndex("�������"))
            str�������� = .TextMatrix(.Row, .ColIndex("��������"))
        End With
    If str������� = "" Then Exit Sub
    Call frmEdit.ShowMe(Me, 1, str��������, lng֧������ID, str֧������, str�������, str��������, blnRefresh)
    If blnRefresh Then Call RefreshData
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_NewItem '����
        Call AddNewInsure
    Case conMenu_Edit_Modify  '�޸�
        Call ModifyInsure
    Case conMenu_Edit_Delete 'ɾ��
        Call DeleteInsure
    Case conMenu_View_Refresh 'ˢ������
        Call RefreshData
    Case Else
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean, blnHaveData As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    With vs֧�����
        blnEnable = Val(.RowData(.Row)) = 0
        blnHaveData = .TextMatrix(.Row, .ColIndex("�������")) <> ""
    End With
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        Control.Enabled = False
    Case conMenu_Edit_NewItem
        Control.Enabled = blnEnable And Not blnHaveData
    Case conMenu_Edit_Modify
        Control.Enabled = blnEnable And blnHaveData
    Case conMenu_Edit_Delete
        Control.Enabled = blnEnable And blnHaveData
    Case Else
    End Select
End Sub

Private Sub vs֧�����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs֧�����, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs֧�����_DblClick()
    Dim str������� As String
    
    With vs֧�����
        If .Row <= 0 Then Exit Sub
         If .RowData(.Row) = 1 Then Exit Sub
        str������� = .TextMatrix(.Row, .ColIndex("�������"))
    End With
    If str������� = "" Then
        Call AddNewInsure
    Else
        Call ModifyInsure
    End If
End Sub

Private Sub vs֧�����_GotFocus()
    If vs֧�����.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs֧�����, &HFFEBD7
End Sub

Private Sub vs֧�����_LostFocus()
    If vs֧�����.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs֧�����
    OS.OpenIme False
End Sub

Private Sub vs֧�����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub
