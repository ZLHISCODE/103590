VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmBlackTypeManage 
   BorderStyle     =   0  'None
   Caption         =   "������Ϊ����"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRuleBack 
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   180
      ScaleHeight     =   3210
      ScaleWidth      =   9825
      TabIndex        =   3
      Top             =   3975
      Width           =   9825
      Begin VSFlex8Ctl.VSFlexGrid vsGridRule 
         Height          =   2700
         Left            =   285
         TabIndex        =   4
         Top             =   285
         Width           =   7035
         _cx             =   12409
         _cy             =   4762
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
         FormatString    =   $"frmBlackTypeManage.frx":0000
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
      Begin VB.Shape shpRule 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   8745
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.PictureBox picTypeBack 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   465
      ScaleHeight     =   3285
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   210
      Width           =   8025
      Begin VSFlex8Ctl.VSFlexGrid vsGridType 
         Height          =   2700
         Left            =   375
         TabIndex        =   2
         Top             =   450
         Width           =   7035
         _cx             =   12409
         _cy             =   4762
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
         FormatString    =   $"frmBlackTypeManage.frx":0075
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
      Begin VB.Shape shpBorder 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   7425
         Top             =   1230
         Width           =   405
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTitle 
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   5895
         _Version        =   589884
         _ExtentX        =   10398
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "��������>������Ϊ���"
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   30
      Top             =   1035
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBlackTypeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String
Private Const conPane_Type = 1
Private Const conPane_Rule = 2

Public Event zlActivate(ByVal frmSubForm As Form) '�¼�����
Public Event zlChangeType() '�ı�����Ϊ���

Public Function zlLoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-13 15:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    On Error GoTo errHandle
    
    Call LoadTypeDataToGrid
    
    With vsGridType
        If Not (.Row > .Rows - 1 Or .Row < 1) Then
            strType = .TextMatrix(.Row, .ColIndex("����"))
        End If
    End With
    Call LoadRuleDataToGrid(strType)
    zlLoadData = True
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
    '����:2018-11-08 17:54:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String
    Dim panThis As Pane
    
    On Error GoTo errHandle
    
    Set panThis = dkpMan.CreatePane(conPane_Type, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "������Ϊ����"
    panThis.Handle = picTypeBack.hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = conPane_Type
    
    Set panThis = dkpMan.CreatePane(conPane_Rule, 250, 580, DockBottomOf, panThis)
    panThis.Title = ""
    panThis.Tag = conPane_Rule
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picRuleBack.hwnd
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Type
        Item.Handle = picTypeBack.hwnd
    Case conPane_Rule
        Item.Handle = picRuleBack.hwnd
    End Select
End Sub
Private Sub InitTypeGridColumnHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������ͷ
    '����:���˺�
    '����:2018-11-08 15:13:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGridType
        .Clear: .Rows = 2: .Cols = 5
        i = 0
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "��Ч����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ϵͳ�̶�": .ColWidth(i) = 1000: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        zl_vsGrid_Para_Restore mlngModule, vsGridType, Me.Caption, "������Ϊ�����б�"
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub InitRuleGridColumHead(ByVal cllType As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ƹ���������ͷ
    '     cllType-ԤԼ��ʽ��
    '���:str��Ϊ���
    '����:���˺�
    '����:2018-11-08 18:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    
    On Error GoTo errHandle
    
    If cllType Is Nothing Then Set cllType = New Collection
    
    With vsGridRule
        .Clear:
        .Rows = 3: .Cols = cllType.Count + 6
        .FixedRows = IIf(cllType.Count = 0, 1, 2)
        .Rows = .FixedRows + 1
        If cllType.Count = 0 Then
            i = 0
            .TextMatrix(0, i) = "���ƹ���": .ColWidth(i) = 2000: i = i + 1
            .TextMatrix(0, i) = "����ԤԼ": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "�Һ�": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "����": .ColWidth(i) = 800: i = i + 1
        Else
            i = 0
            .TextMatrix(0, i) = "���ƹ���"
            .TextMatrix(1, i) = "���ƹ���": .ColWidth(i) = 2000: i = i + 1
            
            .TextMatrix(0, i) = "����ԤԼ":
            .TextMatrix(1, i) = "����ԤԼ": .ColWidth(i) = 800: i = i + 1
            For j = 1 To cllType.Count
                .TextMatrix(0, i) = "ԤԼ��ʽ"
                .TextMatrix(1, i) = cllType(j): .ColWidth(i) = 800: i = i + 1
            Next
            .TextMatrix(0, i) = "�Һ�":
            .TextMatrix(1, i) = "�Һ�": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "��Ժ":
            .TextMatrix(1, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "��Ժ"
            .TextMatrix(1, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
            .TextMatrix(0, i) = "����"
            .TextMatrix(1, i) = "����": .ColWidth(i) = 800: i = i + 1
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(IIf(cllType.Count <> 0, 1, 0), i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) <> "���ƹ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .MergeCol(i) = True
        Next
        .MergeCells = flexMergeRestrictAll
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = cllType.Count <> 0
        .MergeRow(1) = cllType.Count <> 0
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadRuleDataToGrid(ByVal strType As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2018-11-08 16:17:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strName As String, lngRow As Long, cllType As Collection
    Dim i As Long, int���Ʒ�ʽ As Integer
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select a.Ӧ�ó���,a.��Ϊ���,a.ԤԼ��ʽ,a.���,a.���ƹ���,a.���Ʒ�ʽ,b.���� as ԤԼ���� " & _
    "   From ������Ϊ���� A,ԤԼ��ʽ B  " & _
    "   where a.��Ϊ���=[1] and a.ԤԼ��ʽ=b.����(+) " & _
    "   Order by ��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strType)
    
    Set cllType = New Collection
    
    rsTemp.Sort = "ԤԼ����"
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            
            If Nvl(!ԤԼ��ʽ) <> "" Then
                blnFind = False
                For i = 1 To cllType.Count
                    If cllType(i) = Nvl(!ԤԼ��ʽ) Then
                        blnFind = True: Exit For
                    End If
                Next
                If blnFind = False Then
                    cllType.Add Nvl(!ԤԼ��ʽ)
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    vsGridRule.Redraw = flexRDNone
    Call InitRuleGridColumHead(cllType)
    
    rsTemp.Sort = "���"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsGridRule
        .Clear 1
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not rsTemp.EOF
            lngRow = .FindRow(Nvl(rsTemp!���ƹ���), 1, .ColIndex("���ƹ���"))
            If lngRow = -1 Then
                If .TextMatrix(.Rows - 1, .ColIndex("���ƹ���")) <> "" Then .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
            
            .TextMatrix(lngRow, .ColIndex("���ƹ���")) = Nvl(rsTemp!���ƹ���)
            If Nvl(rsTemp!Ӧ�ó���) = "ԤԼ" Then
                
                If Trim(Nvl(rsTemp!ԤԼ��ʽ)) = "" Then
                    .TextMatrix(lngRow, .ColIndex("����ԤԼ")) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
                Else
                    .TextMatrix(lngRow, .ColIndex(Trim(Nvl(rsTemp!ԤԼ��ʽ)))) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
                End If
            ElseIf Nvl(rsTemp!Ӧ�ó���) <> "" Then
                .TextMatrix(lngRow, .ColIndex(rsTemp!Ӧ�ó���)) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
            
            End If
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsGridRule.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadTypeDataToGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2018-11-08 16:17:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lngRow As Long
    Dim strName As String
    Dim lngPreRow As Long, blnFind As Boolean
    
    On Error GoTo errHandle
    
    strSQL = "Select ����,����,����,�Ƿ�̶�,��Ч���� From ������Ϊ���� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsGridType
        If .Row > 0 And .Row <= .Rows - 1 Then
            strName = .TextMatrix(.Row, .ColIndex("����"))
            lngPreRow = .Row
        Else
            lngPreRow = 1
        End If
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Redraw = flexRDNone
        blnFind = False
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��Ч����")) = IIf(Val(Nvl(rsTemp!��Ч����)) <> 0, Val(Nvl(rsTemp!��Ч����)) & "����", "")
            .TextMatrix(lngRow, .ColIndex("�Ƿ�ϵͳ�̶�")) = IIf(Val(Nvl(rsTemp!�Ƿ�̶�)) = 1, "��", "")
            If strName = .TextMatrix(lngRow, .ColIndex("����")) Then .Row = lngRow: blnFind = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        If blnFind = False Then
            If lngPreRow > .Rows - 1 Or lngPreRow < 1 Then
                If .Rows >= 2 Then .Row = 1
            Else
                .Row = lngPreRow
            End If
        End If
        vsGridType_AfterRowColChange 0, 0, .Row, .Col
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӷ���(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ķ���(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "���ƹ������(&S)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 8122
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӷ���", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ķ���", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "���ƹ������", cbrControl.Index + 1): cbrControl.BeginGroup = True
        cbrControl.IconId = 8122
        
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
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
Private Function IsAllowEdit(ByVal lngRow As Long, Optional blnNotCheckSys As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����Ƿ�����༭
    '���:lngRow-ָ����
    '     blnNotCheckSys-�����̶���
    '����:����༭����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
        
    If lngRow <= 0 Or lngRow > vsGridType.Rows - 1 Then Exit Function
    With vsGridType
        If blnNotCheckSys Then
            IsAllowEdit = .TextMatrix(lngRow, .ColIndex("����")) <> ""
        Else
            IsAllowEdit = .TextMatrix(lngRow, .ColIndex("����")) <> "" And .TextMatrix(lngRow, .ColIndex("�Ƿ�ϵͳ�̶�")) = ""
        End If
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
    
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�༭������Ϊ����")
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If vsGridType.Rows >= 2 Then
           Control.Enabled = vsGridType.TextMatrix(1, vsGridType.ColIndent("����")) <> ""
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
        Control.Enabled = Control.Visible And IsAllowEdit(vsGridType.Row, True)
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGridType.Row)
    Case conMenu_Edit_Privacy
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGridType.Row, True)
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
    Case conMenu_Edit_Privacy: Call ExecuteModifyRule
    Case conMenu_View_Refresh: LoadTypeDataToGrid
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ExecuteAddItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�����Ӳ�����Ϊ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListTypeEdit
    On Error GoTo errHandle
    If Not frmEdit.zlShowEdit(mfrmMain, EM_Ty_����) Then Exit Function
    Call LoadTypeDataToGrid
    RaiseEvent zlChangeType
    ExecuteAddItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteModifyItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���޸Ĳ�����Ϊ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListTypeEdit
    Dim strCode As String
    On Error GoTo errHandle
    With vsGridType
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("�Ƿ�ϵͳ�̶�")) <> "" Then
'            MsgBox  "��ǰ���Ϊϵͳ�̶���,��ֻ���޸���Ч���޺Ϳ��ƹ�m!", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    If strCode = "" Then Exit Function
    
    If Not frmEdit.zlShowEdit(mfrmMain, EM_Ty_�޸�, strCode) Then Exit Function
    Call LoadTypeDataToGrid
    RaiseEvent zlChangeType
    ExecuteModifyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteModifyRule() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���޸Ĺ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListTypeEdit
    Dim strCode As String, strName As String
    
    On Error GoTo errHandle
    With vsGridType
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        strCode = Trim(.TextMatrix(.Row, .ColIndex("����")))
        strName = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    If strCode = "" Then Exit Function
    
    If Not frmEdit.zlShowEdit(mfrmMain, EM_Ty_�������, strCode) Then Exit Function
    
    Call LoadRuleDataToGrid(strName)
    ExecuteModifyRule = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ɾ��������Ϊ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 17:10:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCode As String, strName As String, lngRow As Long, strSQL As String
    
    On Error GoTo errHandle
    With vsGridType
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("�Ƿ�ϵͳ�̶�")) <> "" Then
            MsgBox "�������ϵͳ�̶��Ĳ�����Ϊ�������ɾ��!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("����")))
        strName = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    If strCode = "" Then Exit Function
     
    
    If MsgBox("��ȷ��Ҫ�Բ�����Ϊ����Ϊ��" & strName & "������ɾ������ ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strSQL = "Zl_������Ϊ����_Delete('" & strCode & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    With vsGridType
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
    RaiseEvent zlChangeType
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
    If Me.ActiveControl Is Nothing Then vsGridType.SetFocus
    RaiseEvent zlActivate(Me)
End Sub

Private Sub Form_Load()

    Err = 0: On Error GoTo errHandle
    RestoreWinState Me, App.ProductName
    
    Call InitPancel
    Call InitTypeGridColumnHead
    Call LoadTypeDataToGrid
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



 Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zl_vsGrid_Para_Save mlngModule, vsGridType, Me.Caption, "������Ϊ�����б�"
    Err = 0: On Error Resume Next
    Set mcbsMain = Nothing
    Set mfrmMain = Nothing
End Sub
Private Sub picRuleBack_Resize()
    Err = 0: On Error Resume Next
    With picRuleBack
        shpRule.Top = 0
        shpRule.Left = 0
        shpRule.Width = .ScaleWidth
        shpRule.Height = .ScaleHeight
        
        vsGridRule.Left = 10: vsGridRule.Top = 10
        vsGridRule.Width = .ScaleWidth - vsGridRule.Left
        vsGridRule.Height = .ScaleHeight - vsGridRule.Top - 10
    End With
End Sub

Private Sub picTypeBack_Resize()
    Err = 0: On Error Resume Next
    With picTypeBack
        stcTitle.Move 6, 6, .ScaleWidth
        shpBorder.Left = 0
        shpBorder.Top = 0
        shpBorder.Height = .ScaleHeight - shpBorder.Top
        shpBorder.Width = .ScaleWidth
        
        vsGridType.Left = 10: vsGridType.Top = stcTitle.Top + stcTitle.Height + 10
        
        vsGridType.Width = .ScaleWidth - vsGridType.Left
        vsGridType.Height = .ScaleHeight - vsGridType.Top - 20
    End With
End Sub

Private Sub vsGridType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    Dim strCode As String
    Dim strName As String
  
    With vsGridType
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        If OldRow = NewRow Then Exit Sub
        
        strCode = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        strName = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        Call LoadRuleDataToGrid(strName)
    End With
End Sub

Private Sub vsGridType_DblClick()
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
    Dim blnRule As Boolean, strType As String
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo errHandle
    blnRule = Me.ActiveControl Is vsGridRule
    If blnRule Then
        With vsGridType
            If .Row > 0 And .Row <= .Rows - 1 Then
                strType = .TextMatrix(.Row, .ColIndex("����"))
            End If
        End With
    End If
    objOut.Title.Text = IIf(blnRule, strType & "���ƹ����嵥", "���ò�����Ϊ�嵥")
    Set objOut.Body = IIf(blnRule, vsGridRule, vsGridType)
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
    If vsGridType.Visible Then vsGridType.SetFocus
End Sub

Private Sub vsGridType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
