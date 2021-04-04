VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicWorkTimeManage 
   BorderStyle     =   0  'None
   Caption         =   "�ϰ�ʱ�����"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsfWorkTime 
      Height          =   2955
      Left            =   870
      TabIndex        =   0
      Top             =   1110
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicWorkTimeManage.frx":0000
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
      Left            =   180
      Top             =   120
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   630
      TabIndex        =   1
      Top             =   540
      Width           =   7905
      _Version        =   589884
      _ExtentX        =   13944
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "��������>�ϰ�ʱ�����"
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
Attribute VB_Name = "frmClinicWorkTimeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mGridHead
    COL_վ�� = 0
    COL_����
    Col_ʱ���
    COL_�ϰ�ʱ��
    COL_��Ϣʱ��
    COL_����Ԥ��ʱ��
    COL_ȱʡԤԼʱ��
    COL_��ǰ�Һ�ʱ��
End Enum

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar

    Err = 0: On Error GoTo ErrHandler
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����ʱ���(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ʱ���(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ʱ���(&D)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "ˢ������(&B)", cbrControl.Index)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����ʱ���", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�ʱ���", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��ʱ���", cbrControl.index + 1)
        .Item(cbrControl.index + 1).BeginGroup = True
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnabled As Boolean, blnVisible As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnEnabled = Not vsfWorkTime.IsSubtotal(vsfWorkTime.Row) And vsfWorkTime.Rows > 1
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "ʱ�������")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfWorkTime.Rows > vsfWorkTime.FixedRows
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem '����
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '�޸�
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_Delete 'ɾ��
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strվ�� As String, str���� As String, strʱ��� As String
    Dim frmEdit As frmClinicWorkTimeEdit
    
    Err = 0: On Error GoTo ErrHandler
    
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem '�����ϰ�ʱ��
        Set frmEdit = New frmClinicWorkTimeEdit
        If frmEdit.ShowMe(Me, Fun_Add) Then Call LoadData: Set grsWorkTime = Nothing '����ȡ�ϰ�ʱ��
    Case conMenu_Edit_Modify '�����ϰ�ʱ��
        With vsfWorkTime
            If .Row <= 0 Then Exit Sub
            If .IsSubtotal(.Row) Then Exit Sub
            
            strվ�� = .Cell(flexcpData, .Row, COL_վ��)
            str���� = Trim(.TextMatrix(.Row, COL_����))
            strʱ��� = .TextMatrix(.Row, Col_ʱ���)
            Set frmEdit = New frmClinicWorkTimeEdit
            If frmEdit.ShowMe(Me, Fun_Update, strվ��, str����, strʱ���) Then Call LoadData: Set grsWorkTime = Nothing '����ȡ�ϰ�ʱ��
        End With
    Case conMenu_Edit_Delete 'ɾ���ϰ�ʱ��
        If ExcuteDelete() Then Call LoadData: Set grsWorkTime = Nothing '����ȡ�ϰ�ʱ��
    Case conMenu_View_Refresh 'ˢ������
        Call LoadData
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExcuteDelete() As Boolean
    '����:ִ��ɾ������
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim strվ�� As String, str���� As String, strʱ��� As String
    Dim blnUsed As Boolean
    On Error GoTo errHandle
    
    With vsfWorkTime
        If .Row <= 0 Then Exit Function
        If .IsSubtotal(.Row) Then Exit Function
        
        strվ�� = .Cell(flexcpData, .Row, COL_վ��)
        str���� = Trim(.TextMatrix(.Row, COL_����))
        strʱ��� = .TextMatrix(.Row, Col_ʱ���)
    End With
    
    strSQL = "Select 1 From �ٴ������Դ���� Where �ϰ�ʱ�� = [1] And Rownum < 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 1 From �ٴ��������� Where �ϰ�ʱ�� = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strʱ���)
    If rsTemp.EOF Then
        If MsgBox("��ȷ��Ҫɾ���ϰ�ʱ��Σ�" & strʱ��� & "����", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        blnUsed = True
        If MsgBox("ע�⣺" & vbCrLf & _
                  "    �ϰ�ʱ��Σ�" & strʱ��� & "�������ѱ�ʹ�ã�ɾ��������Ҫ������ʹ���˸��ϰ�ʱ����������˷�ʱ�εİ��Ž������»���ʱ�Σ����򣬿��ܻᵼ��ԤԼ�Һų���" & vbCrLf & _
                  vbCrLf & _
                  "    ��ȷ��Ҫɾ���ϰ�ʱ��Σ�" & strʱ��� & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    'ɾ����Ч�Լ��
    If CheckHaveUsed(strվ��, str����, strʱ���) Then
        MsgBox "��ǰ�ϰ�ʱ����ѱ�ʹ�ã�����ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    'Zl_�ϰ�ʱ��_Delete(
    strSQL = "Zl_�ϰ�ʱ��_Delete("
    'վ��_In   ʱ���.վ��%Type,
    strSQL = strSQL & "'" & strվ�� & "',"
    '����_In   ʱ���.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    'ʱ���_In ʱ���.ʱ���%Type
    strSQL = strSQL & "'" & strʱ��� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If blnUsed Then
        MsgBox "ע�⣺" & vbCrLf & _
               "    �ϰ�ʱ��Σ�" & strʱ��� & "���ѱ�ɾ�����뼰ʱ������ʹ���˸��ϰ�ʱ����������˷�ʱ�εİ��Ž������»���ʱ�Σ����򣬿��ܻᵼ��ԤԼ�Һų���", vbExclamation, gstrSysName
    End If
    
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckHaveUsed(ByVal strվ�� As String, ByVal str���� As String, ByVal strʱ��� As String) As Boolean
    '��鵱ǰ�ϰ�ʱ����Ƿ��ѱ�ʹ��
    Dim strSQL As String, rs�ϰ�ʱ�� As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    '���ԭ�ϰ�ʱ���Ƿ�ʹ�ã���ʹ�õĲ����޸�վ�㡢���ࡢʱ���
    '����ɾ����ʹ�õķ�Χ������һ��,��ʹ�õ�ʱ��ֻҪ��һ�����ɣ���ͬվ�㣬��ͬ������ܻ��ж��ͬ����ʱ��Σ�
    '�ٴ������Դ����
    strSQL = "Select 1" & vbNewLine & _
            " From (Select b.�ϰ�ʱ��, c.վ��, a.����," & vbNewLine & _
            "              Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.�ϰ�ʱ��, c.վ�� Desc, a.���� Desc) As ���" & vbNewLine & _
            "        From �ٴ������Դ A, �ٴ������Դ���� B, ���ű� C" & vbNewLine & _
            "        Where a.Id = b.��Դid And a.����id = c.Id)" & vbNewLine & _
            " Where ��� = 1 And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-') And �ϰ�ʱ�� = [3] And Rownum < 2"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    If Not rs�ϰ�ʱ�� Is Nothing Then
        If Not rs�ϰ�ʱ��.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ���������(�̶�����ģ��)
    strSQL = "Select 1" & vbNewLine & _
            " From (Select a.�ϰ�ʱ��, c.վ��, b.����," & vbNewLine & _
            "              Row_Number() Over(Partition By a.�ϰ�ʱ�� Order By a.�ϰ�ʱ��, c.վ�� Desc, b.���� Desc) As ���" & vbNewLine & _
            "        From �ٴ��������� A, �ٴ����ﰲ�� D, �ٴ������Դ B, ���ű� C" & vbNewLine & _
            "        Where a.����id = d.Id And d.��Դid = b.Id And b.����id = c.Id)" & vbNewLine & _
            " Where ��� = 1 And Nvl(վ��, '-') = Nvl([1], '-') And Nvl(����, '-') = Nvl([2], '-') And �ϰ�ʱ�� = [3] And Rownum < 2"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strվ��, str����, strʱ���)
    If Not rs�ϰ�ʱ�� Is Nothing Then
        If Not rs�ϰ�ʱ��.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '�ٴ������¼
    '����飬��Ϊ�ñ�̫������ϰ�ʱ�ε���Ϣ����������������У�û���ҵ��ϰ�ʱ��ʱ������������������ȡ
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    Call InitGrid
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitGrid()
    Dim strHead As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strHead = "վ��,4,0|����,4,1000|ʱ���,4,1000|�ϰ�ʱ��,4,1500|��Ϣʱ��,4,1500|����Ԥ��ʱ��,4,1200|" & _
            "ȱʡԤԼʱ��,4,1200|��ǰ�Һ�ʱ��,4,1200"
    With vsfWorkTime
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AutoSizeMode = flexAutoSizeRowHeight
        .RowHeightMin = 300
        .WordWrap = True
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        Call RestoreFlexState(vsfWorkTime, App.ProductName & "\" & Me.Name)
        .Redraw = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadData() As Boolean
    '�����ϰ�ʱ�������
    Dim i As Long, lngRow As Long, strSQL As String
    Dim rs�ϰ�ʱ�� As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    strSQL = "Select a.ʱ���, a.����, a.��Ϣʱ��, a.��ʼʱ��, a.��ֹʱ��," & vbNewLine & _
            "        a.ȱʡʱ��, a.��ǰʱ��, a.����Ԥ��ʱ��," & vbNewLine & _
            "        b.���, b.���� As վ��" & vbNewLine & _
            " From ʱ��� A, Zlnodelist B" & vbNewLine & _
            " Where a.վ�� = b.���(+)" & vbNewLine & _
            " Order By Nvl(b.���, -1), Nvl(a.����, -1)"
    Set rs�ϰ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ϰ�ʱ���")
    With vsfWorkTime
        lngRow = .Row
        .Redraw = False
        .Clear 1
        .Subtotal flexSTClear
        .Rows = rs�ϰ�ʱ��.RecordCount + 1
        i = 1
        Do While Not rs�ϰ�ʱ��.EOF
            .TextMatrix(i, COL_վ��) = Nvl(rs�ϰ�ʱ��!վ��, "ȫԺ")
            .Cell(flexcpData, i, COL_վ��) = Nvl(rs�ϰ�ʱ��!���)
            .TextMatrix(i, COL_����) = Nvl(rs�ϰ�ʱ��!����, " ")
            .TextMatrix(i, Col_ʱ���) = Nvl(rs�ϰ�ʱ��!ʱ���)
            .TextMatrix(i, COL_�ϰ�ʱ��) = Format(Nvl(rs�ϰ�ʱ��!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rs�ϰ�ʱ��!��ֹʱ��), "hh:mm")
            .TextMatrix(i, COL_��Ϣʱ��) = FormatStr(Nvl(rs�ϰ�ʱ��!��Ϣʱ��))
            .TextMatrix(i, COL_����Ԥ��ʱ��) = IIf(Val(Nvl(rs�ϰ�ʱ��!����Ԥ��ʱ��)) = 0, "", Nvl(rs�ϰ�ʱ��!����Ԥ��ʱ��))
            .TextMatrix(i, COL_ȱʡԤԼʱ��) = Format(Nvl(rs�ϰ�ʱ��!ȱʡʱ��), "hh:mm")
            .TextMatrix(i, COL_��ǰ�Һ�ʱ��) = Format(Nvl(rs�ϰ�ʱ��!��ǰʱ��), "hh:mm")
            .RowData(i) = IIf(i Mod 2 = 0, vbWindowBackground, G_AlternateColor) '���������н�����ɫ
            i = i + 1
            rs�ϰ�ʱ��.MoveNext
        Loop
        .AutoSize 0, .Cols - 1
        
        '�����н�����ɫ
        For i = 1 To .Rows - 1
            .Cell(flexcpBackColor, i, Col_ʱ���, i, .Cols - 1) = .RowData(i)
        Next
        
        Call DataSplitGroup '������ʾ
        If .Rows > 1 Then 'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            If lngRow = 0 Then
                .Row = 1
            ElseIf lngRow > .Rows - 1 Then
                .Row = .Rows - 1
            Else
                .Row = lngRow
            End If
        End If
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatStr(ByVal strIn As String) As String
    '��ʽ����ʼʱ��1-��ֹʱ��1; ��ʼʱ��2-��ֹʱ��2;��.��
    Dim varRow As Variant, varCol As Variant
    Dim i As Integer
    Dim strReturn As String
    
    If strIn = "" Then FormatStr = "": Exit Function
    varRow = Split(strIn, ";")
    For i = 0 To UBound(varRow)
        strReturn = strReturn & vbCrLf
        varCol = Split(varRow(i), "-")
        strReturn = strReturn & Format(varCol(0), "hh:mm") & "-" & Format(varCol(1), "hh:mm")
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, 3)
    FormatStr = strReturn
End Function

Private Sub DataSplitGroup()
    Dim i As Integer, j As Integer

    Err = 0: On Error GoTo ErrHandler
    With vsfWorkTime
        .OutlineBar = flexOutlineBarComplete '����/������ʾĿ¼��������
        .OutlineCol = COL_���� '���������
        .Outline COL_����
        
        .Subtotal flexSTClear
        .Subtotal flexSTNone, COL_վ��, , , , , True, "%s", , True
        .SubtotalPosition = flexSTAbove
        .MergeCells = flexMergeRestrictRows
        .MergeCol(COL_����) = True

        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then '�Ƿ���С��
                .MergeRow(i) = True
                .IsCollapsed(i) = flexOutlineExpanded '�Ƿ�չ��״̬
                .Cell(flexcpText, i, 1, i, .Cols - 1) = .Cell(flexcpTextDisplay, i, 0) 'Flexcptextdisplay ��Ԫ���ʽ���˵��ı�����(ֻ��)
                .RowHeight(i) = 300
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = flexAlignLeftCenter
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 0, 0, Me.ScaleWidth
    With vsfWorkTime
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(vsfWorkTime, App.ProductName & "\" & Me.Name)
End Sub


Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfWorkTime.Visible And vsfWorkTime.Enabled Then vsfWorkTime.SetFocus
End Sub

Private Sub vsfWorkTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    '����ѡ������ɫ
    Call SetVsGridRowChangeBackColor(vsfWorkTime, OldRow, NewRow, OldCol, NewCol, _
        vsfWorkTime.BackColorSel, Col_ʱ���, vsfWorkTime.Cols - 1)
End Sub

Private Sub vsfWorkTime_DblClick()
    Dim blnUpdate As Boolean
    Dim strվ�� As String, str���� As String, strʱ��� As String
    Dim frmEdit As frmClinicWorkTimeEdit
    
    Err = 0: On Error GoTo ErrHandler
    With vsfWorkTime
        If .Row < 1 Then Exit Sub
        If .IsSubtotal(.Row) Then Exit Sub
        
        strվ�� = .Cell(flexcpData, .Row, COL_վ��)
        str���� = Trim(.TextMatrix(.Row, COL_����))
        strʱ��� = .TextMatrix(.Row, Col_ʱ���)
        
        Set frmEdit = New frmClinicWorkTimeEdit
        If zlStr.IsHavePrivs(mstrPrivs, "ʱ�������") Then
            '�޸�
            If frmEdit.ShowMe(Me, Fun_Update, strվ��, str����, strʱ���) Then
                Call LoadData 'ˢ������
                Set grsWorkTime = Nothing '����ȡ�ϰ�ʱ��
            End If
        Else
            '�鿴
            frmEdit.ShowMe Me, Fun_View, strվ��, str����, strʱ���
        End If
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfWorkTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
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

Private Sub zlDataPrint(BytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    objOut.Title.Text = "�ϰ�ʱ����嵥"
    Set objOut.Body = vsfWorkTime
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If BytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, BytMode
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
