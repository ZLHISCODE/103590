VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBlackListReasonManage 
   BorderStyle     =   0  'None
   Caption         =   "���ò�����Ϊԭ��"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2955
      Left            =   420
      TabIndex        =   1
      Top             =   645
      Width           =   7035
      _cx             =   12409
      _cy             =   5212
      Appearance      =   0
      BorderStyle     =   0
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
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
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
      FormatString    =   $"frmBlackListReasonManage.frx":0000
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
   Begin XtremeSuiteControls.ShortcutCaption stcTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _Version        =   589884
      _ExtentX        =   10398
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "��������>������Ϊ����ԭ��"
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
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   0
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmBlackListReasonManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String
Public Event zlActivate(ByVal frmSubForm As Form) '�¼�����

Public Sub zlInitComm(frmMain As Form, cbsThis As Object, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿ�
    '���:objPati-����������
    '     cbsThis-�˵�����
    '     strPrivs-Ȩ�޴�
    '     lngModule-ģ���
    '����:���˺�
    '����:2018-11-08 11:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain: Set mcbsMain = cbsThis
    mstrPrivs = strPrivs: mlngModule = lngModule
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Sub zlCancelBands()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ؼ����
    '����:���˺�
    '����:2018-11-15 15:48:53
    '��Ҫ�����ؽ�ǰ��ɾ���ؼ��󣬿��ܴ��ڰ󶨵Ŀؼ����ڹ�������������У����ɾ��ʱ������ؼ�һ��ɾ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrToolBar As CommandBar
    On Error GoTo errHandle
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "������")
    If cbrToolBar Is Nothing Then Exit Sub
    cbrToolBar.Controls.DeleteAll
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Function zlLoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-13 15:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    On Error GoTo errHandle
    zlLoadData = LoadDataToGrid
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGridColumnHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������ͷ
    '����:���˺�
    '����:2018-11-08 15:13:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGrid
        .Clear: .Rows = 2: .Cols = 4
        i = 0
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ϵͳ�̶�": .ColWidth(i) = 1000: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "����ԭ���б�"
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Function LoadDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2018-11-08 16:17:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lngRow As Long
    Dim strName As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ����,����,����,�Ƿ�̶� From ���ò�����Ϊԭ�� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsGrid
        If .Row > 0 And .Row <= .Rows - 1 Then
            strName = .TextMatrix(.Row, .ColIndex("����"))
        End If
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�Ƿ�ϵͳ�̶�")) = IIf(Val(Nvl(rsTemp!�Ƿ�̶�)) = 1, "��", "")
            If strName = .TextMatrix(lngRow, .ColIndex("����")) Then .Row = lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadDataToGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHandle
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    
    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӳ���ԭ��(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ĳ���ԭ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������ԭ��(&D)")
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
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "������")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = mcbsMain.Add("������", xtpBarTop)
    End If
    
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup And cbrControl.Index > 1 Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӳ���ԭ��", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ĳ���ԭ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������ԭ��", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
'    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "�����ҹ��ˡ�")
'    objPopup.flags = xtpFlagRightAlign
'    '���󶨵Ŀؼ����붯̬���أ���Ϊ������һ����ɾ�������󶨵Ŀؼ��ľ���ͻ���0
'    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
'
'    If txtFind.UBound > 0 Then Unload txtFind(1)
'    Load txtFind(1)
'    objCustom.Handle = txtFind(1).hWnd
'    objCustom.flags = xtpFlagRightAlign
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function IsAllowEdit(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����Ƿ�����༭
    '���:lngRow-ָ����
    '����:����༭����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
        
    If lngRow <= 0 Or lngRow > vsGrid.Rows - 1 Then Exit Function
    With vsGrid
        IsAllowEdit = .TextMatrix(lngRow, .ColIndex("����")) <> "" And .TextMatrix(lngRow, .ColIndex("�Ƿ�ϵͳ�̶�")) = ""
    End With
    Exit Function
errHandle:
    Exit Function
End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù��ܲ˵���Eanbled���Ժ�visible����
    '����:���˺�
    '����:2018-11-08 16:55:37
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim blnVisible As Boolean, blnEnable As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�༭����ԭ��")
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If vsGrid.Rows >= 2 Then
           Control.Enabled = vsGrid.TextMatrix(1, vsGrid.ColIndent("����")) <> ""
        Else
           Control.Enabled = False
        End If
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGrid.Row)
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGrid.Row)
    End Select
End Sub
Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ع��ܲ���
    '����:���˺�
    '����:2018-11-08 16:56:26
    '---------------------------------------------------------------------------------------------------------------------------------------------

      
    Err = 0: On Error GoTo errHandle
    
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem: Call ExecuteAddItem
    Case conMenu_Edit_Modify: Call ExecuteModifyItem
    Case conMenu_Edit_Delete: Call ExcuteDelete
    Case conMenu_View_Refresh: LoadDataToGrid
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ExecuteAddItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�����ӳ���ԭ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListReasonEdit
    On Error GoTo errHandle
    If Not frmEdit.zlShowEdit(mfrmMain, 0) Then Exit Function
    Call LoadDataToGrid
    ExecuteAddItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteModifyItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���޸ĳ���ԭ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListReasonEdit
    Dim strCode As String
    On Error GoTo errHandle
    With vsGrid
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("�Ƿ�ϵͳ�̶�")) <> "" Then
            MsgBox "�������ϵͳ�̶��ĳ���ԭ������޸�!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    If strCode = "" Then Exit Function
    
    If Not frmEdit.zlShowEdit(mfrmMain, 1, strCode) Then Exit Function
    Call LoadDataToGrid
    ExecuteModifyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ɾ������ԭ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 17:10:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCode As String, strName As String, lngRow As Long
    Dim strSQL As String
    
    On Error GoTo errHandle
    With vsGrid
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("�Ƿ�ϵͳ�̶�")) <> "" Then
            MsgBox "�������ϵͳ�̶��ĳ���ԭ�����ɾ��!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("����")))
        strName = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    If strCode = "" Then Exit Function
     
    
    If MsgBox("��ȷ��Ҫ�Գ���ԭ��Ϊ��" & strName & "������ɾ������ ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strSQL = "Zl_���ò�����Ϊԭ��_Delete('" & strCode & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    With vsGrid
        lngRow = .Row
        If lngRow > .Rows - 1 And .Rows <= 2 Then
            .Clear 1: .Rows = 2
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        ElseIf lngRow > .Rows - 1 Then
            .RemoveItem lngRow
            .Row = .Rows - 1
        ElseIf lngRow <= .Rows - 1 Then
            .RemoveItem lngRow
            .Row = lngRow - 1
        End If
    End With
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
     
Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then vsGrid.SetFocus
    RaiseEvent zlActivate(Me)
End Sub

Private Sub Form_Load()

    Err = 0: On Error GoTo errHandle
    RestoreWinState Me, App.ProductName
    
    Call InitGridColumnHead
    Call LoadDataToGrid
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    stcTitle.Move 0, 0, Me.ScaleWidth
    With vsGrid
        .Left = 10: .Top = stcTitle.Top + stcTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "����ԭ���б�"
    Err = 0: On Error Resume Next
    Set mcbsMain = Nothing
    Set mfrmMain = Nothing
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "����ԭ���б�"
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "����ԭ���б�"
End Sub

Private Sub vsGrid_DblClick()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:˫���޸�
    '����:���˺�
    '����:2018-11-08 17:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ExecuteModifyItem
End Sub

 

Private Sub zlDataPrint(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2018-11-08 17:37:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo errHandle
    objOut.Title.Text = "���ò�����Ϊԭ���嵥"
    Set objOut.Body = vsGrid
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub stcTitle_GotFocus()
    On Error Resume Next
    If vsGrid.Visible Then vsGrid.SetFocus
End Sub


Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandle
    If Not (Button = vbRightButton) Or Not (Me.Visible And Me.Enabled) Then Exit Sub
    
    Me.SetFocus:   RaiseEvent zlActivate(Me)
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
