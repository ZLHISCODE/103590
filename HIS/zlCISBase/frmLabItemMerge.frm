VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemMerge 
   BorderStyle     =   0  'None
   Caption         =   "���������Ŀ����"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   3960
      ScaleHeight     =   3675
      ScaleWidth      =   3855
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkUpper 
         Caption         =   "���ִ�Сд(&U)"
         Height          =   210
         Left            =   2055
         TabIndex        =   10
         Top             =   0
         Width           =   2040
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��"
         Height          =   350
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   2235
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   ">"
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   1260
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "<"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   900
         Width           =   390
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   3495
         Picture         =   "frmLabItemMerge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   225
         Width           =   360
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   435
         TabIndex        =   2
         Top             =   225
         Width           =   3045
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��"
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   1860
         Width           =   390
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   3120
         Left            =   435
         TabIndex        =   4
         Top             =   555
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   5503
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
         Caption         =   "������Ŀ:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   0
         Width           =   810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3825
      _cx             =   6747
      _cy             =   6482
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
      Cols            =   5
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
End
Attribute VB_Name = "frmLabItemMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid
Private mbln��� As Boolean         '��ǰ��Ŀ�Ƿ������Ŀ
Private mstr���� As String          '��ǰ��Ŀ�ļ�������

Private Enum mCol
    ID = 0: ���: ����: ������: Ӣ����
End Enum

Dim ObjItem As ListItem
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
            .Rows = 1: .FixedRows = 1: .Cols = 5: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.���) = "���": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.������) = "������": .TextMatrix(0, mCol.Ӣ����) = "Ӣ����"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.���) = 500: .ColWidth(mCol.����) = 1000
        .ColWidth(mCol.������) = 2000: .ColWidth(mCol.Ӣ����) = 2000
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColAlignment(mCol.���) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.���) = lngCount
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
        
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    mbln��� = False: mstr���� = ""
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select �����Ŀ, ��������, ����Ӧ�� From ������ĿĿ¼ Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        If .RecordCount > 0 Then
            mbln��� = (Val("" & !�����Ŀ) = 1)
            mstr���� = "" & !��������
'            If Val("" & !����Ӧ��) <> 1 Then
'                zlRefresh = False: Exit Function
'            End If
        Else
            zlRefresh = False: Exit Function
        End If
    End With
    
    gstrSql = "Select Distinct R.������Ŀid As ID, R.������� As ���, V.����, V.������, V.Ӣ����" & vbNewLine & _
            " From ���鱨����Ŀ R, ����������Ŀ V, ����ϲ����� H , ������ĿĿ¼ I " & vbNewLine & _
            " Where R.������Ŀid = V.ID And R.������Ŀid =H.�ϲ���ĿID And H.����ĿID=[1] And ϸ��id Is Null" & vbNewLine & _
            " and r.������ĿID  = i.id and i.�����Ŀ <> 1 and ����Ӧ�� = 1"
    gstrSql = gstrSql & " Union ALL " & _
            " Select Distinct  R.������Ŀid As ID, 0 As ���, i.����, i.����, '' As Ӣ����" & vbNewLine & _
            " From ���鱨����Ŀ R, ����������Ŀ V, ����ϲ����� H, ������ĿĿ¼ I" & vbNewLine & _
            " Where R.������Ŀid = V.ID And R.������Ŀid =H.�ϲ���ĿID And H.����ĿID=[1] And ϸ��id Is Null" & vbNewLine & _
            " and r.������ĿID  = i.id and i.�����Ŀ = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngItemId-ָ���༭����Ŀ
        
    Me.Tag = "�༭": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strLists As String
    
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '���ݱ���
    gstrSql = "Zl_����ϲ�����_Edit(" & mlngItemID & ",'" & strLists & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long
    With Me.vfgList
        Select Case Index
        Case 0         '���
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set ObjItem = Me.lvwItem.SelectedItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mCol.ID) = Mid(ObjItem.Key, 2)
            .TextMatrix(.Rows - 1, mCol.����) = ObjItem.Text
            .TextMatrix(.Rows - 1, mCol.������) = ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1)
            .TextMatrix(.Rows - 1, mCol.Ӣ����) = ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1)
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwItem.ListItems.Remove ObjItem.Key: Me.lvwItem.SetFocus
        Case 1          'ɾ��
            If .Row < .FixedRows Then Exit Sub
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.����))
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = .TextMatrix(.Row, mCol.������)
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = .TextMatrix(.Row, mCol.Ӣ����)
            ObjItem.Selected = True
            .RemoveItem .Row
        Case 2          '����
            If .Row <= .FixedRows Then Exit Sub
            lngCurRow = .Row
            .AddItem "", lngCurRow - 1
            .TextMatrix(lngCurRow - 1, mCol.ID) = .TextMatrix(lngCurRow + 1, mCol.ID)
            .TextMatrix(lngCurRow - 1, mCol.����) = .TextMatrix(lngCurRow + 1, mCol.����)
            .TextMatrix(lngCurRow - 1, mCol.������) = .TextMatrix(lngCurRow + 1, mCol.������)
            .TextMatrix(lngCurRow - 1, mCol.Ӣ����) = .TextMatrix(lngCurRow + 1, mCol.Ӣ����)
            .RemoveItem lngCurRow + 1
            .Row = lngCurRow - 1
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
            
        Case 3          '����
            If .Row >= .Rows - 1 Then Exit Sub
            lngCurRow = .Row
            .AddItem "", lngCurRow
            .TextMatrix(lngCurRow, mCol.ID) = .TextMatrix(lngCurRow + 2, mCol.ID)
            .TextMatrix(lngCurRow, mCol.����) = .TextMatrix(lngCurRow + 2, mCol.����)
            .TextMatrix(lngCurRow, mCol.������) = .TextMatrix(lngCurRow + 2, mCol.������)
            .TextMatrix(lngCurRow, mCol.Ӣ����) = .TextMatrix(lngCurRow + 2, mCol.Ӣ����)
            .RemoveItem lngCurRow + 2
            .Row = lngCurRow + 1
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        End Select
        
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.���) = lngCount
        Next
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String
    
    If Me.chkUpper.Value = 0 Then
        strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
        gstrSql = "Select distinct I.ID, I.����, I.���� As ������, L.��д As Ӣ����, Nvl(H.����ĿID,0) as ʹ��" & vbNewLine & _
                "From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L, ����ϲ����� H " & vbNewLine & _
                "Where I.ID=H.�ϲ���ĿID(+) And I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ <> 1 and i.����Ӧ�� = 1 And I.�������� = '" & mstr���� & "' And" & vbNewLine & _
                "      (I.���� Like '" & strFind & "%' Or Upper(I.����) Like '" & gstrMatch & strFind & "%' Or Upper(L.��д) Like '" & gstrMatch & strFind & "%')"
        gstrSql = gstrSql & " Union ALL " & _
                " Select distinct I.ID, I.����, I.���� As ������, '' As Ӣ����, Nvl(H.����ĿID,0) as ʹ��" & vbNewLine & _
                " From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L, ����ϲ����� H " & vbNewLine & _
                " Where I.ID=H.�ϲ���ĿID(+) And I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ = 1 And I.�������� = '" & mstr���� & "' And" & vbNewLine & _
                "      (I.���� Like '" & strFind & "%' Or Upper(I.����) Like '" & gstrMatch & strFind & "%' Or Upper(L.��д) Like '" & gstrMatch & strFind & "%')"
    Else
        strFind = DelInvalidChar(Trim(Me.txtFind.Text))
        gstrSql = "Select distinct I.ID, I.����, I.���� As ������, L.��д As Ӣ����, Nvl(H.����ĿID,0) as ʹ��" & vbNewLine & _
                "From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L, ����ϲ����� H" & vbNewLine & _
                "Where I.ID=H.�ϲ���ĿID(+) And I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ <> 1 and i.����Ӧ�� =1 And I.�������� = '" & mstr���� & "' And" & vbNewLine & _
                "      (I.���� Like '" & strFind & "%' Or I.���� Like '" & gstrMatch & strFind & "%' Or L.��д Like '" & gstrMatch & strFind & "%')"
        gstrSql = gstrSql & " Union ALL " & _
                " Select distinct I.ID, I.����, I.���� As ������, '' As Ӣ����, Nvl(H.����ĿID,0) as ʹ��" & vbNewLine & _
                " From ������ĿĿ¼ I, ���鱨����Ŀ R, ������Ŀ L, ����ϲ����� H" & vbNewLine & _
                " Where I.ID=H.�ϲ���ĿID(+) And I.ID = R.������Ŀid And R.������Ŀid = L.������Ŀid And I.�����Ŀ = 1 And I.�������� = '" & mstr���� & "' And" & vbNewLine & _
                "      (I.���� Like '" & strFind & "%' Or I.���� Like '" & gstrMatch & strFind & "%' Or L.��д Like '" & gstrMatch & strFind & "%')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        .Filter = " ʹ�� = 0 "
        Do While Not .EOF
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = "" & !������
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = "" & !Ӣ����
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    
    '����ѡ����
    Me.lvwItem.ListItems.Remove "_" & mlngItemID
    
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
        .Add , "_������", "������", 2300
        .Add , "_Ӣ����", "Ӣ����", 600
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 180
    Me.picEdit.Height = Me.ScaleHeight - Me.picEdit.Top - 180
    If Me.Tag = "�༭" Then
        Me.vfgList.Width = Me.picEdit.Left - Me.vfgList.Left
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
    Else
        Me.vfgList.Width = Me.picEdit.Left + Me.picEdit.Width - Me.vfgList.Left
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
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
