VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemCost 
   BorderStyle     =   0  'None
   Caption         =   "��Ŀ�����Լ�"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7710
      _cx             =   13600
      _cy             =   3334
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   7710
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2295
      Width           =   7710
      Begin VB.CheckBox chkHand 
         Caption         =   "�����ֹ���(&M)"
         Height          =   195
         Left            =   6075
         TabIndex        =   11
         Top             =   135
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ѡ������(&1)"
         Height          =   180
         Index           =   1
         Left            =   6090
         TabIndex        =   10
         Top             =   1995
         Width           =   1560
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ѡ���Լ�(&0)"
         Height          =   180
         Index           =   0
         Left            =   6090
         TabIndex        =   9
         Top             =   1680
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ���Լ������б���ɾ��"
         Height          =   350
         Index           =   1
         Left            =   2595
         TabIndex        =   6
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ��ӵ��Լ������б���"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "���ҡ�    "
         Height          =   350
         Left            =   6075
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   6075
         TabIndex        =   3
         Top             =   720
         Width           =   1605
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2280
         Left            =   0
         TabIndex        =   2
         Top             =   450
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6075
         TabIndex        =   7
         Top             =   495
         Width           =   810
      End
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " ����˵�����õ�λ���������¸����Լ�����ɱ���Ŀ��������˴Ρ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   345
      TabIndex        =   8
      Top             =   120
      Width           =   5490
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   120
      Picture         =   "frmLabItemCost.frx":0000
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "frmLabItemCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ�ļ�����Ŀ��������Ŀid
Private mlngLabID As Long          '��ǰ��ʾ�ļ�����Ŀ��������Ŀid

Private Enum mCol
    ID = 0: ����: ����: ��λ: �ֹ�
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = mCol.�ֹ� + 1
            .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.����) = "����"
            .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.��λ) = "��λ"
            .TextMatrix(0, mCol.�ֹ�) = "�ֹ�"
        End If
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.����) = 1000
        .ColWidth(mCol.����) = 2500: .ColWidth(mCol.��λ) = 500
        If Me.chkHand.Value = 0 Then
            .ColWidth(mCol.�ֹ�) = 0
        Else
            .ColWidth(mCol.�ֹ�) = 500
        End If
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .AutoSize mCol.�ֹ�, .Cols - 1
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '���ܣ���������idˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    Dim rsApt As New ADODB.Recordset, strColSql As String
    
    mlngItemID = lngItemID
    mlngLabID = 0
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
    
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    On Error GoTo ErrHand
    gstrSql = "Select R.������Ŀid From ���鱨����Ŀ R Where R.������Ŀid = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    If rsTemp.RecordCount > 0 Then
        mlngLabID = Nvl(rsTemp!������ĿID, 0)
'    Else
'        MsgBox "�ü�����Ŀ��Ϣ���ֶ�ʧ��", vbInformation, gstrSysName
    End If
    
    gstrSql = "Select Distinct A.ID, A.����, A.����" & vbNewLine & _
            "From �����Լ���ϵ L, �������� A" & vbNewLine & _
            "Where L.����id = A.ID And L.��Ŀid = [1]" & vbNewLine & _
            "Order By A.����"
    Set rsApt = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngLabID)
    With rsApt
        strColSql = ""
        Do While Not .EOF
            strColSql = strColSql & ",Max(Decode(L.����id," & !ID & ",L.����)) As C" & !ID
            .MoveNext
        Loop
    End With
    
    gstrSql = "Select L.����id As ID, I.����, I.����, I.���㵥λ As ��λ, Max(Decode(L.����id, Null, L.����)) As �ֹ�" & strColSql & vbNewLine & _
            "From ������ĿĿ¼ I, �������� T, �����Լ���ϵ L" & vbNewLine & _
            "Where L.����id = T.����id And T.����id = I.ID And L.��Ŀid = [1]" & vbNewLine & _
            "Group By L.����id, I.����, I.����, I.���㵥λ"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngLabID)
    
    Me.vfgList.FixedCols = 0
    Set Me.vfgList.DataSource = rsTemp
    With rsApt
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Me.vfgList.ColData(.AbsolutePosition + mCol.�ֹ�) = CLng(!ID)
            Me.vfgList.Cell(flexcpData, 0, .AbsolutePosition + mCol.�ֹ�) = CStr(!����)
            Me.vfgList.TextMatrix(0, .AbsolutePosition + mCol.�ֹ�) = "" & !����
            .MoveNext
        Loop
    End With
    Me.vfgList.FixedCols = mCol.�ֹ�
    
    Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngItemId-ָ���༭����Ŀ
    Me.Tag = "�༭": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strLists As String, lngRow As Long, lngCol As Long
    
    strLists = ""
    With Me.vfgList
        For lngRow = .FixedRows To .Rows - 1
            If Me.chkHand.Value = vbChecked Then
                If Val(.TextMatrix(lngRow, mCol.�ֹ�)) > 99999999 Then
                    MsgBox "��" & .TextMatrix(lngRow, mCol.����) & "��(" & lngRow & "��)�����ֹ������˴�����̫��", vbInformation, gstrSysName
                    .Row = lngRow: .Col = mCol.�ֹ�: .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Val(.TextMatrix(lngRow, mCol.�ֹ�)) <> 0 Then
                    strLists = strLists & "|" & .TextMatrix(lngRow, mCol.ID) & ";;" & Val(.TextMatrix(lngRow, mCol.�ֹ�))
                End If
            End If
            For lngCol = mCol.�ֹ� + 1 To .Cols - 1
                If Val(.TextMatrix(lngRow, lngCol)) > 99999999 Then
                    MsgBox "��" & .TextMatrix(lngRow, mCol.����) & "��(" & lngRow & "��)����" & .TextMatrix(0, lngCol) & "��(" & lngCol & "��)�˴�����̫��", vbInformation, gstrSysName
                    .Row = lngRow: .Col = lngCol: .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then
                    strLists = strLists & "|" & .TextMatrix(lngRow, mCol.ID) & ";" & .ColData(lngCol) & ";" & Val(.TextMatrix(lngRow, lngCol))
                End If
            Next
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '���ݱ���
    gstrSql = "Zl_�����Լ���ϵ_Edit(" & mlngLabID & ",'" & strLists & "')"
    If LenB(gstrSql) > 4000 Then
        MsgBox "������̫����Լ�����Ʒ�����ܱ��棡", vbInformation, gstrSysName
        Me.vfgList.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0
End Function

Private Sub chkKind_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkUpper_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkHand_Click()
    With Me.vfgList
        If Me.chkHand.Value = 0 Then
            .ColWidth(mCol.�ֹ�) = 0: .ColHidden(mCol.�ֹ�) = True
        Else
            .ColWidth(mCol.�ֹ�) = 500: .ColHidden(mCol.�ֹ�) = False
        End If
    End With
End Sub

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCol As Long
    With Me.vfgList
        If Me.opt����(0).Value Then
            Select Case Index
            Case 0         '���
                If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
                Set objItem = Me.lvwItem.SelectedItem
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, mCol.ID) = Mid(objItem.Key, 2)
                .TextMatrix(.Rows - 1, mCol.����) = objItem.Text
                .TextMatrix(.Rows - 1, mCol.����) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1)
                .TextMatrix(.Rows - 1, mCol.��λ) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1)
                If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
                Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
            Case 1          'ɾ��
                If .Row < .FixedRows Then MsgBox "�����Լ����Ѿ�ɾ����", vbInformation, gstrSysName: Exit Sub
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.����))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = .TextMatrix(.Row, mCol.����)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1) = .TextMatrix(.Row, mCol.��λ)
                objItem.Selected = True
                .RemoveItem .Row
            End Select
        Else
            Select Case Index
            Case 0         '���
                If .Cols >= mCol.�ֹ� + 6 Then MsgBox "���ֻ������6��������", vbInformation, gstrSysName: Exit Sub
                If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
                Set objItem = Me.lvwItem.SelectedItem
                .Cols = .Cols + 1
                .ColData(.Cols - 1) = Val(Mid(objItem.Key, 2))
                .Cell(flexcpData, 0, .Cols - 1) = objItem.Text
                .TextMatrix(0, .Cols - 1) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1)
                .Col = .Cols - 1
                If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
                Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
            Case 1          'ɾ��
                If .Col <= mCol.�ֹ� Then MsgBox "��ѡ����ɾ���������У�", vbInformation, gstrSysName: Exit Sub
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .ColData(.Col), .Cell(flexcpData, 0, .Col))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = .TextMatrix(0, .Col)
                objItem.Selected = True
                If .Col < .Cols - 1 Then
                    For lngCol = .Col To .Cols - 2
                        .ColData(lngCol) = .ColData(lngCol + 1)
                        .Cell(flexcpData, 0, lngCol) = .Cell(flexcpData, 0, lngCol + 1)
                        .TextMatrix(0, lngCol) = .TextMatrix(0, lngCol + 1)
                    Next
                End If
                .Cols = .Cols - 1
            End Select
            .AutoSize mCol.�ֹ�, .Cols - 1
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String, strColdId As String
    
    strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
    If Me.opt����(0).Value Then
        gstrSql = "Select Distinct M.����id As ID, I.����, I.����, I.���㵥λ As ��λ" & vbNewLine & _
                "From ������ĿĿ¼ I, ������Ŀ���� N, �������� M" & vbNewLine & _
                "Where I.ID = N.������Ŀid And I.ID = M.����id And" & vbNewLine & _
                "      (I.���� Like '" & strFind & "%' Or N.���� Like '" & gstrMatch & strFind & "%' Or N.���� Like '" & gstrMatch & strFind & "%')"
    Else
        gstrSql = "Select I.ID, I.����, I.����, '' As ��λ" & vbNewLine & _
                "From �������� I" & vbNewLine & _
                "Where I.���� Like '" & strFind & "%' Or I.���� Like '" & gstrMatch & strFind & "%' Or I.���� Like '" & gstrMatch & strFind & "%'"
    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = "" & !����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1) = "" & !��λ
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgList
        If Me.opt����(0).Value Then
            For lngCount = .FixedRows To .Rows - 1
                Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mCol.ID)
            Next
        Else
            For lngCount = mCol.�ֹ� + 1 To .Cols - 1
                strColdId = .ColData(lngCount)
                Me.lvwItem.ListItems.Remove "_" & strColdId
            Next
        End If
    End With
    
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "û��ƥ�����Ŀ��", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgList.SetFocus
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_����", "����", 1000
        .Add , "_����", "����", 3500
        .Add , "_��λ", "��λ", 1000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 105
    If Me.Tag = "�༭" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 105
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub opt����_Click(Index As Integer)
    Me.lvwItem.ListItems.Clear
    Me.txtFind.Text = "": Me.txtFind.SetFocus
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.lvwItem.Height = Me.picEdit.ScaleHeight - Me.lvwItem.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "�༭" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col < mCol.�ֹ� Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < mCol.�ֹ� Then Cancel = True: Exit Sub
    If Row < Me.vfgList.FixedRows Then Cancel = True
End Sub


