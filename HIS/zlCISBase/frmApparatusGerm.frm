VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApparatusGerm 
   BorderStyle     =   0  'None
   Caption         =   "����ϸ��ͨ��"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton optWay 
      Caption         =   "����������ͨ��(&2)"
      Height          =   180
      Index           =   1
      Left            =   1980
      TabIndex        =   10
      Top             =   120
      Width           =   2085
   End
   Begin VB.OptionButton optWay 
      Caption         =   "����ϸ��ͨ��(&1)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Value           =   -1  'True
      Width           =   1800
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   135
      ScaleHeight     =   2505
      ScaleWidth      =   8145
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2685
      Width           =   8145
      Begin VB.CheckBox chkUpper 
         Caption         =   "���ִ�Сд(&U)"
         Height          =   210
         Left            =   6060
         TabIndex        =   7
         Top             =   1605
         Width           =   1755
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ������ͨ���б���ɾ��"
         Height          =   350
         Index           =   1
         Left            =   2610
         TabIndex        =   6
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ��ӵ�����ͨ���б���"
         Height          =   350
         Index           =   0
         Left            =   15
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "���ҡ�    "
         Height          =   350
         Left            =   6060
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   6060
         TabIndex        =   1
         Top             =   720
         Width           =   1755
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   450
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   3625
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
         Left            =   6060
         TabIndex        =   0
         Top             =   495
         Width           =   810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2265
      Left            =   135
      TabIndex        =   4
      Top             =   390
      Width           =   8145
      _cx             =   14367
      _cy             =   3995
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
End
Attribute VB_Name = "frmApparatusGerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long          '��ǰ��ʾ������id

Private Enum mcol
    ID = 0: ���: ����: ������: Ӣ����: ͨ����
End Enum

Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
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
            .Rows = 1: .FixedRows = 1: .Cols = 6: .FixedCols = 0
        End If
        .TextMatrix(0, mcol.ID) = "ID": .TextMatrix(0, mcol.���) = "���": .TextMatrix(0, mcol.����) = "����"
        .TextMatrix(0, mcol.������) = "������": .TextMatrix(0, mcol.Ӣ����) = "Ӣ����": .TextMatrix(0, mcol.ͨ����) = "ͨ����"
        
        .ColWidth(mcol.ID) = 0: .ColWidth(mcol.���) = 450: .ColWidth(mcol.����) = 700
        .ColWidth(mcol.������) = 2500: .ColWidth(mcol.Ӣ����) = 2500: .ColWidth(mcol.ͨ����) = 720
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColAlignment(mcol.���) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mcol.���) = lngCount
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngAptId As Long) As Boolean
    '���ܣ���������idˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim rsTemp As New ADODB.Recordset
    
    mLngAptId = lngAptId
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
    
    If lngAptId = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    If Me.optWay(0).Value Then
        gstrSql = "Select G.ID, Rownum As ���, G.����, G.������, G.Ӣ����, L.ͨ������ As ͨ����" & vbNewLine & _
                "From ����ϸ������ L, ����ϸ�� G" & vbNewLine & _
                "Where L.ϸ��id = G.ID And L.����id = [1]"
    Else
        gstrSql = "Select G.ID, Rownum As ���, G.����, G.������, G.Ӣ����, L.ͨ������ As ͨ����" & vbNewLine & _
                "From ����ϸ������ L, �����ÿ����� G" & vbNewLine & _
                "Where L.������id = G.ID And L.����id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False
End Function

Public Function zlEditStart() As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ lngAptId-ָ���༭����Ŀ
        
    Me.Tag = "�༭": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mLngAptId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim strLists As String, strItems As String
    
    Err = 0: On Error GoTo ErrHand
    strLists = ""
    With Me.vfgList
        If .Rows = 1 Then
            If Me.optWay(0).Value Then
                gstrSql = "Zl_����ϸ������_Edit(0," & mLngAptId & ",0,Null,0)"
            Else
                gstrSql = "Zl_����ϸ������_Edit(1," & mLngAptId & ",0,Null,0)"
            End If
            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        End If
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mcol.ID)) = 0 Then
                MsgBox "��" & lngCount & "����Ŀ��ȷ����", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If Trim(.TextMatrix(lngCount, mcol.ͨ����)) = "" Then
                MsgBox "��" & lngCount & "�С�ͨ���롱δ��д��", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(.TextMatrix(lngCount, mcol.ͨ����)), vbFromUnicode)) > 50 Then
                MsgBox "��" & lngCount & "�С�ͨ���롱��������(50���ַ�)��", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            strItems = .TextMatrix(lngCount, mcol.ID)
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mcol.ͨ����))
            strLists = strLists & "|" & strItems

            '���ݱ���
            If Me.optWay(0).Value Then
                gstrSql = "Zl_����ϸ������_Edit(0," & mLngAptId & "," & .TextMatrix(lngCount, mcol.ID) & _
                          ",'" & Trim(.TextMatrix(lngCount, mcol.ͨ����)) & "'," & lngCount & ")"
            Else
                gstrSql = "Zl_����ϸ������_Edit(1," & mLngAptId & "," & .TextMatrix(lngCount, mcol.ID) & _
                          ",'" & Trim(.TextMatrix(lngCount, mcol.ͨ����)) & "'," & lngCount & ")"
            End If

            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        Next
    End With
'    If LenB(gstrSql) > 4000 Then
'        MsgBox "ͨ����Ŀ����̫�࣬���ܱ��棡", vbInformation, gstrSysName
'        Me.vfgList.SetFocus: zlEditSave = 0: Exit Function
'    End If

'    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mLngAptId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub chkUpper_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long
    With Me.vfgList
        Select Case Index
        Case 0         '���
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwItem.SelectedItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mcol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Rows - 1, mcol.����) = objItem.Text
            .TextMatrix(.Rows - 1, mcol.������) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1)
            .TextMatrix(.Rows - 1, mcol.Ӣ����) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1)
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
        Case 1          'ɾ��
            If .Row < .FixedRows Then Exit Sub
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mcol.ID), .TextMatrix(.Row, mcol.����))
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = .TextMatrix(.Row, mcol.������)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = .TextMatrix(.Row, mcol.Ӣ����)
            objItem.Selected = True
            .RemoveItem .Row
        End Select
        
        For lngCount = .Row To .Rows - 1
            .TextMatrix(lngCount, mcol.���) = lngCount
        Next
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String, strTable As String
    If Me.optWay(0).Value Then
        strTable = "����ϸ��"
    Else
        strTable = "�����ÿ�����"
    End If
    
    
    If Me.chkUpper.Value Then
        strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
        gstrSql = "Select ID, ����, ������, Ӣ����" & vbNewLine & _
                "From " & strTable & vbNewLine & _
                "Where ���� Like '" & strFind & "%' Or Upper(������) Like '" & gstrMatch & strFind & "%' Or" & vbNewLine & _
                "      Upper(Ӣ����) Like '" & gstrMatch & strFind & "%' Or Upper(����) Like '" & gstrMatch & strFind & "%'"
    Else
        strFind = DelInvalidChar(Trim(Me.txtFind.Text))
        gstrSql = "Select ID, ����, ������, Ӣ����" & vbNewLine & _
                "From " & strTable & vbNewLine & _
                "Where ���� Like '" & strFind & "%' Or ������ Like '" & gstrMatch & strFind & "%' Or" & vbNewLine & _
                "      Ӣ���� Like '" & gstrMatch & strFind & "%' Or ���� Like '" & gstrMatch & strFind & "%'"
    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_������").Index - 1) = "" & !������
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӣ����").Index - 1) = "" & !Ӣ����
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
'    With Me.vfgList
'        For lngCount = .FixedRows To .Rows - 1
'            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mcol.ID)
'        Next
'    End With
    
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
        .Add , "_����", "����", 700
        .Add , "_������", "������", 2500
        .Add , "_Ӣ����", "Ӣ����", 2500
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
        Me.optWay(1).Enabled = False: Me.optWay(0).Enabled = False
    Else
        Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 105
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
        Me.optWay(1).Enabled = True: Me.optWay(0).Enabled = True
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

Private Sub optWay_Click(Index As Integer)
    Call Me.zlRefresh(mLngAptId)
'    If Me.Tag <> "�༭" Then Exit Sub
'    Debug.Print "ss"
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.lvwItem.Height = Me.picEdit.ScaleHeight - Me.lvwItem.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "�༭" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(1, "|;'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mcol.ͨ���� Then Cancel = True
    If Row < Me.vfgList.FixedRows Then Cancel = True
End Sub
