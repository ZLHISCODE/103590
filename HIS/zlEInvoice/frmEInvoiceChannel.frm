VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceChannel 
   BorderStyle     =   0  'None
   Caption         =   "�շ���������"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vs�շ����� 
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceChannel.frx":0000
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
      Begin VB.Image imgAdd 
         Height          =   240
         Left            =   -480
         Picture         =   "frmEInvoiceChannel.frx":0123
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
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
      Caption         =   "�շ���������"
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   360
      Top             =   2760
      Width           =   525
   End
End
Attribute VB_Name = "frmEInvoiceChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs�շ�����.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Public Sub RefreshData()
    Dim strSql As String, i As Integer
    Dim rs�շ����� As ADODB.Recordset

    On Error GoTo errHandle
    strSql = "Select a.���� As ���㷽ʽ, a.���� As ��������, c.Id As �����id, Nvl(c.����, '��') As ���������," & vbNewLine & _
                "      b.��������, 0 As ���ӱ�־" & vbNewLine & _
                "From ���㷽ʽ A, �շ��������� B, ҽ�ƿ���� C" & vbNewLine & _
                "Where a.���� = b.���㷽ʽ(+) And a.���� = c.���㷽ʽ(+) And b.�����id Is Null" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.���㷽ʽ, c.���� As ��������, a.�����id, Nvl(b.����, '��') As ���������, a.��������," & vbNewLine & _
                "       Decode(a.���㷽ʽ, b.���㷽ʽ, 0, 1) As ���ӱ�־" & vbNewLine & _
                "From �շ��������� A, ҽ�ƿ���� B, ���㷽ʽ C" & vbNewLine & _
                "Where a.���㷽ʽ = c.���� And a.�����id = b.Id And c.���� = 8" & vbNewLine & _
                "Order By �����id, ���ӱ�־"

    Set rs�շ����� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    With vs�շ�����
        .Clear 1
        .Rows = 2
        For i = 1 To rs�շ�����.RecordCount
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = rs�շ�����!���㷽ʽ
            .TextMatrix(i, .ColIndex("ԭ���㷽ʽ")) = rs�շ�����!���㷽ʽ
            .TextMatrix(i, .ColIndex("��������")) = Val(rs�շ�����!��������)
            .TextMatrix(i, .ColIndex("�����ID")) = Val(NVL(rs�շ�����!�����id))
            If Val(rs�շ�����!��������) = 8 And Val(NVL(rs�շ�����!�����id)) > 0 And Val(NVL(rs�շ�����!���ӱ�־)) = 0 Then
                .CellButtonPicture = imgAdd: .ComboList = "..."
            End If
            .TextMatrix(i, .ColIndex("���������")) = NVL(rs�շ�����!���������)
            .TextMatrix(i, .ColIndex("��������")) = NVL(rs�շ�����!��������)
            .TextMatrix(i, .ColIndex("���ӱ�־")) = NVL(rs�շ�����!���ӱ�־)
            rs�շ�����.MoveNext
            If i < rs�շ�����.RecordCount Then .Rows = .Rows + 1
        Next
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

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    blnEnable = vs�շ�����.TextMatrix(vs�շ�����.Row, vs�շ�����.ColIndex("��������")) <> ""

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        Control.Enabled = False
    Case conMenu_Edit_NewItem
        Control.Enabled = Not blnEnable
    Case conMenu_Edit_Modify
        Control.Enabled = blnEnable
    Case conMenu_Edit_Delete
        Control.Enabled = blnEnable
    Case Else
    End Select
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

Private Sub vs�շ�����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With vs�շ�����
        If .ColIndex("���㷽ʽ") <> Col Then Exit Sub
        If .TextMatrix(Row, .ColIndex("���㷽ʽ")) = "" Then Exit Sub
        If .TextMatrix(Row, .ColIndex("��������")) = "" Then Exit Sub
        If Val(.TextMatrix(Row, .ColIndex("�����id"))) = 0 Then Exit Sub
        If Modify�շ���������(.TextMatrix(Row, .ColIndex("���㷽ʽ")), Val(.TextMatrix(Row, .ColIndex("�����id"))), .TextMatrix(Row, .ColIndex("��������")), _
           .TextMatrix(Row, .ColIndex("ԭ���㷽ʽ"))) = False Then Exit Sub
        Call RefreshData
    End With
End Sub

Private Sub vs�շ�����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim str���㷽ʽ As String
    Dim rs���������� As ADODB.Recordset
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs�շ�����, OldRow, NewRow, OldCol, NewCol
    
    With vs�շ�����
        If NewCol <> .ColIndex("���㷽ʽ") Then
            .ComboList = "..."
            Exit Sub
        End If
        If Val(.TextMatrix(NewRow, .ColIndex("��������"))) = 8 And Val(.TextMatrix(NewRow, .ColIndex("�����id"))) > 0 Then
            If Val(.TextMatrix(NewRow, .ColIndex("���ӱ�־"))) = 0 Then
                .CellButtonPicture = imgAdd: .ComboList = "..."
                .ColComboList(.ColIndex("���㷽ʽ")) = ""
            Else
                str���㷽ʽ = .TextMatrix(NewRow - 1, .ColIndex("���㷽ʽ"))
                Set rs���������� = Get�������㷽ʽ(str���㷽ʽ)
                .ColComboList(.ColIndex("���㷽ʽ")) = .BuildComboList(rs����������, "���㷽ʽ", "���㷽ʽ")
            End If
        End If
    End With
End Sub

Private Sub vs�շ�����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str���㷽ʽ As String
    Dim rs����������  As New ADODB.Recordset
    With vs�շ�����
        Select Case Col
        Case .ColIndex("���㷽ʽ")
            If Not (Val(.TextMatrix(Row, .ColIndex("��������"))) = 8 And Val(.TextMatrix(Row, .ColIndex("�����id"))) > 0) Then Cancel = True: Exit Sub
            If Val(.TextMatrix(Row, .ColIndex("���ӱ�־"))) = 1 Then
                str���㷽ʽ = .TextMatrix(Row - 1, .ColIndex("���㷽ʽ"))
                If str���㷽ʽ <> "" Then
                    Set rs���������� = Get�������㷽ʽ(str���㷽ʽ)
                    .ColComboList(.ColIndex("���㷽ʽ")) = .BuildComboList(rs����������, "���㷽ʽ", "���㷽ʽ")
                    Exit Sub
                End If
            End If
            If CheckThirdCard(Val(.TextMatrix(Row, .ColIndex("�����id")))) = False Then
                Cancel = True: Exit Sub
            Else
                .CellButtonPicture = imgAdd: .ComboList = "..."
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vs�շ�����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vs�շ�����
        If .ColIndex("���㷽ʽ") <> Col Then Exit Sub
        If Not (Val(.TextMatrix(Row, .ColIndex("��������"))) = 8 And Val(.TextMatrix(Row, .ColIndex("�����id"))) > 0) Then Exit Sub
        If CheckThirdCard(Val(.TextMatrix(Row, .ColIndex("�����id")))) = False Then Exit Sub
        .AddItem Val(.TextMatrix(Row, .ColIndex("��������"))) & vbTab & .TextMatrix(Row, .ColIndex("�����id")) & vbTab & .TextMatrix(Row, .ColIndex("���������")) & vbTab & .TextMatrix(Row, .ColIndex("ԭ���㷽ʽ")) & vbTab & "" & vbTab & "" & vbTab & "1", Row + 1
    End With
End Sub

Private Sub vs�շ�����_DblClick()
    Dim str�������� As String
    With vs�շ�����
        If .Row <= 0 Then Exit Sub
        str�������� = .TextMatrix(.Row, .ColIndex("��������"))
    End With
    If str�������� = "" Then
        Call AddNewChannel
    Else
        Call ModifyChannel
    End If
End Sub

Private Sub vs�շ�����_GotFocus()
    If vs�շ�����.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs�շ�����, &HFFEBD7
End Sub

Private Sub vs�շ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    With vs�շ�����
        If .Row < 1 Then Exit Sub
        If .Col <> .ColIndex("���㷽ʽ") Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("��������")) = "" And Val(.TextMatrix(.Row, .ColIndex("���ӱ�־"))) = 1 Then
            If KeyCode = vbKeyDelete Then .RemoveItem .Row
        End If
    End With
End Sub

Private Sub vs�շ�����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs�շ�����
        Select Case Col
        Case .ColIndex("��������")
            If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End Select
    End With
End Sub

Private Sub vs�շ�����_LostFocus()
    If vs�շ�����.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs�շ�����
    OS.OpenIme False
End Sub

Private Sub vs�շ�����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub

Private Function Get�������㷽ʽ(ByVal str���㷽ʽ As String) As ADODB.Recordset
    Dim strSql As String
    Dim rsTmp  As ADODB.Recordset '��ǰ��Ч��������֧����ʽ
    If str���㷽ʽ = "" Then Exit Function
    On Error GoTo ErrHandler

    strSql = " Select Rownum As ���, a.����, a.���� As ���㷽ʽ, b.���� As ������, Decode(Nvl(b.����, '-'), '-', 1, 0) As ���ӱ�־ " & _
                   " From ���㷽ʽ A, ҽ�ƿ���� B " & _
                   " Where a.���� = 8 And a.���� = b.���㷽ʽ(+) and a.����<>[1]"
    Set Get�������㷽ʽ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str���㷽ʽ)

    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_NewItem '����
        Call AddNewChannel
    Case conMenu_Edit_Modify  '�޸�
        Call ModifyChannel
    Case conMenu_Edit_Delete 'ɾ��
        Call DeleteChannel
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

Private Function CheckThirdCard(ByVal lng�����id As Long) As Boolean
    '���ܣ�����Ƿ��ظ�¼��������
    Dim i As Integer
    
    If lng�����id = 0 Then Exit Function
    With vs�շ�����
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�����id")) = lng�����id And Val(.TextMatrix(i, .ColIndex("���ӱ�־"))) = 1 Then
                Exit Function
            End If
        Next
    End With
    CheckThirdCard = True
End Function

Public Sub AddNewChannel()
    '�����վݷ�Ŀ����
    Dim frmEdit As New frmEInvoiceChannelSet
    Dim blnRefresh As Boolean
    Dim str���㷽ʽ As String, str��������� As String
    Dim lng�����id As Long
    
    On Error GoTo errHandle
    With vs�շ�����
        str���㷽ʽ = .TextMatrix(.Row, .ColIndex("���㷽ʽ"))
        str��������� = .TextMatrix(.Row, .ColIndex("���������"))
        lng�����id = Val(.TextMatrix(.Row, .ColIndex("�����id")))
    End With
    If str���㷽ʽ = "" Then
        MsgBox "δѡ����㷽ʽ������ѡ����㷽ʽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    Call frmEdit.ShowMe(Me, 0, lng�����id, str���������, str���㷽ʽ, , blnRefresh)
    If blnRefresh Then Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteChannel()
    'ɾ���վݷ�Ŀ����
    On Error GoTo errHandle
    Dim lng�����id As Long, str�������� As String
    Dim str���㷽ʽ As String, strSql As String
    Dim str����� As String
    With vs�շ�����
        If .Row = 0 Then Exit Sub
        lng�����id = Val(.TextMatrix(.Row, .ColIndex("�����ID")))
        str���㷽ʽ = .TextMatrix(.Row, .ColIndex("���㷽ʽ"))
        str�������� = .TextMatrix(.Row, .ColIndex("��������"))
        str����� = .TextMatrix(.Row, .ColIndex("���������"))
        
        If str�������� = "" Then Exit Sub
        If MsgBox("��ȷ��Ҫɾ�����㷽ʽΪ��" & str���㷽ʽ & "��" & IIf(lng�����id = 0, "", "���������Ϊ��" & str����� & "��") & "���շ�����������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Me.MousePointer = 11
        strSql = "Zl_�շ���������_Update(2,'" & str���㷽ʽ & "'," & IIf(lng�����id = 0, "NULL", lng�����id) & ",'" & str�������� & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
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

Public Sub ModifyChannel()
    '�޸��վݷ�Ŀ����
    Dim frmEdit As New frmEInvoiceChannelSet
    Dim blnRefresh As Boolean
    Dim str���㷽ʽ As String, str��������� As String
    Dim lng�����id As Long, str�������� As String
    
    On Error Resume Next
    With vs�շ�����
        If .Row = 0 Then Exit Sub
        str���㷽ʽ = .TextMatrix(.Row, .ColIndex("���㷽ʽ"))
        str��������� = .TextMatrix(.Row, .ColIndex("���������"))
        lng�����id = Val(.TextMatrix(.Row, .ColIndex("�����id")))
        str�������� = .TextMatrix(.Row, .ColIndex("��������"))
    End With
    If str�������� = "" Then Exit Sub
    Call frmEdit.ShowMe(Me, 1, lng�����id, str���������, str���㷽ʽ, str��������, blnRefresh)
    If blnRefresh Then Call RefreshData
End Sub

Private Function Modify�շ���������(ByVal str���㷽ʽ As String, ByVal lng�����id As Long, _
                           ByVal str�������� As String, ByVal strԭ���㷽ʽ As String) As Boolean
    Dim strSql As String
    
    If str���㷽ʽ = strԭ���㷽ʽ Then Exit Function
    If lng�����id = 0 Then Exit Function
    If str�������� = "" Then Exit Function
    If strԭ���㷽ʽ = "" Then Exit Function
    
    On Error GoTo errHandle
    '�����շ��������յĽ��㷽ʽ
    strSql = "Zl_�շ���������_Update("
    '��������_In In Number,
    strSql = strSql & 3 & ","
    '���㷽ʽ_In In �շ���������.���㷽ʽ%Type,
    strSql = strSql & "'" & str���㷽ʽ & "',"
    '�����id_In In �շ���������.�����id%Type,
    strSql = strSql & lng�����id & ","
    '��������_In In �շ���������.��������%Type
    strSql = strSql & "'" & str�������� & "',"
    'ԭ���㷽ʽ_In In �շ���������.���㷽ʽ%Type := Null
    strSql = strSql & "'" & strԭ���㷽ʽ & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "�շ���������")
    
    Modify�շ��������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
