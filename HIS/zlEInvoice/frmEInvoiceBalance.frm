VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceBalance 
   BorderStyle     =   0  'None
   Caption         =   "��Ʊ�������"
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
   Begin VSFlex8Ctl.VSFlexGrid vs��Ʊ���� 
      Height          =   1080
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4035
      _cx             =   7117
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceBalance.frx":0000
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   0
      Top             =   2160
      Width           =   525
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "��Ʊ��������"
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
Attribute VB_Name = "frmEInvoiceBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            Call RefreshData
    End Select
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs��Ʊ����.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub vs��Ʊ����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs��Ʊ����
        If .TextMatrix(Row, .ColIndex("���㷽ʽ")) = "" Then Exit Sub
        If .TextMatrix(Row, .ColIndex("��Ʊ���㷽ʽ")) = "" Then Exit Sub
        If Save��Ʊ�������(.TextMatrix(Row, .ColIndex("���㷽ʽ")), .TextMatrix(Row, .ColIndex("��Ʊ���㷽ʽ"))) = False Then Exit Sub
        Call RefreshData
    End With
End Sub

Private Sub vs��Ʊ����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs��Ʊ����, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs��Ʊ����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs��Ʊ����
        If Col <> .ColIndex("��Ʊ���㷽ʽ") Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vs��Ʊ����_GotFocus()
    If vs��Ʊ����.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs��Ʊ����, &HFFEBD7
End Sub

Private Sub vs��Ʊ����_LostFocus()
    If vs��Ʊ����.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs��Ʊ����
    OS.OpenIme False
End Sub

Private Sub vs��Ʊ����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub

Public Sub RefreshData()
    '���ܣ�ˢ������
    Dim strSql As String, i As Integer
    Dim rs��Ʊ���� As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select b.���� as ���㷽ʽ, a.��Ʊ���㷽ʽ From ��Ʊ������� A, ���㷽ʽ B Where a.���㷽ʽ(+) = b.���� And b.���� In (3, 4)"
    Set rs��Ʊ���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With vs��Ʊ����
        .Clear 1
        .Rows = 2
        If rs��Ʊ����.EOF Then Exit Sub
        For i = 1 To rs��Ʊ����.RecordCount
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = rs��Ʊ����!���㷽ʽ
            .TextMatrix(i, .ColIndex("��Ʊ���㷽ʽ")) = Nvl(rs��Ʊ����!��Ʊ���㷽ʽ)
            rs��Ʊ����.MoveNext
            If i < rs��Ʊ����.RecordCount Then .Rows = .Rows + 1
        Next
        .ColComboList(.ColIndex("��Ʊ���㷽ʽ")) = "�����˻�֧��|ҽ��ͳ�����֧��|����ҽ��֧��"
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&E)", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnDelete As Boolean
    On Error Resume Next
    If Not Me.Visible Then Exit Sub

    blnDelete = vs��Ʊ����.TextMatrix(vs��Ʊ����.Row, vs��Ʊ����.ColIndex("��Ʊ���㷽ʽ")) <> ""

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        Control.Enabled = False
    Case conMenu_Edit_Delete
        Control.Enabled = blnDelete
    Case Else
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_Delete 'ɾ��
        Call Delete��Ʊ�������
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

Private Function Save��Ʊ�������(ByVal str���㷽ʽ As String, ByVal str��Ʊ���㷽ʽ As String) As Boolean
    Dim strSql As String
    
    If str���㷽ʽ = "" Then Exit Function
    If str��Ʊ���㷽ʽ = "" Then Exit Function
    
    On Error GoTo errHandle
    'Zl_��Ʊ�������_Update
    strSql = "Zl_��Ʊ�������_Update("
    '��������_In In Number,
    strSql = strSql & 0 & ","
    '���㷽ʽ_In In �շ���������.���㷽ʽ%Type,
    strSql = strSql & "'" & str���㷽ʽ & "',"
    '��Ʊ���㷽ʽ_In In ��Ʊ�������.��Ʊ���㷽ʽ%Type := Null
    strSql = strSql & "'" & str��Ʊ���㷽ʽ & "')"
    
    Call zlDatabase.ExecuteProcedure(strSql, "��Ʊ�������")
    
    Save��Ʊ������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Delete��Ʊ�������()
    Dim strSql As String, str���㷽ʽ As String
    
    With vs��Ʊ����
        If .Row = 0 Then Exit Sub
        str���㷽ʽ = .TextMatrix(.Row, .ColIndex("���㷽ʽ"))
    End With
    If str���㷽ʽ = "" Then Exit Sub
    
    On Error GoTo errHandle
    'Zl_��Ʊ�������_Update
    strSql = "Zl_��Ʊ�������_Update("
    '��������_In In Number,
    strSql = strSql & 1 & ","
    '���㷽ʽ_In In �շ���������.���㷽ʽ%Type,
    strSql = strSql & "'" & str���㷽ʽ & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "��Ʊ�������")

    Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
