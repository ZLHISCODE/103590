VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMediContrast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ѡ����ҩƷ"
   ClientHeight    =   8415
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frmMediContrast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic��ʾ 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   9255
      TabIndex        =   9
      Top             =   120
      Width           =   9255
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMediContrast.frx":6852
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   8925
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10335
      TabIndex        =   4
      Top             =   7800
      Width           =   10335
      Begin VB.CommandButton cmdOk 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   7560
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8760
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000003&
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9840
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9975
      _cx             =   17595
      _cy             =   5741
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6914
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
   Begin VSFlex8Ctl.VSFlexGrid vsf���� 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   9975
      _cx             =   17595
      _cy             =   5741
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6A4F
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
   Begin VSFlex8Ctl.VSFlexGrid vsf���ݻ��� 
      Height          =   2055
      Left            =   10800
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   5775
      _cx             =   10186
      _cy             =   3625
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6B44
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmMediContrast.frx":6C6D
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMediContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long '��¼ѡ��¼����ҩƷID
Private mlng��ѡҩƷID As Long
Private mlng����ҩƷID As Long
Private Const mlngBorderColor As Long = &H0&    'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ
Private Const mcstEditColor = &H80000003   '�ܱ༭����ɫ
Private Const mcstBachColor = &H80000008          '�����˶���ҩƷ����ɫ
Private Const mcstNotBachColor = &H8080FF      'û�����ö���ҩƷ����ɫ
Private mrsFindName As ADODB.Recordset '��ѯ�����ݼ�
Public Sub ShowMe(ByVal objFra As frmMediLists)
    Me.Show vbModal, objFra
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Long, n As Long
    Dim int��� As Integer
    Dim str�������� As String
    
    If vsfList.Rows < 2 Then Exit Sub
    
    With vsf���ݻ���
        For i = 1 To vsfList.Rows - 1
            int��� = 0
            For n = 1 To .Rows - 1
                If Val(vsfList.TextMatrix(i, vsfList.ColIndex("ҩƷID"))) = Val(.TextMatrix(n, .ColIndex("��ѡҩƷID"))) Then
                    int��� = int��� + 1
                    str�������� = int��� & "^" & .TextMatrix(n, .ColIndex("��ѡҩƷID")) & _
                         "^" & .TextMatrix(n, .ColIndex("����ҩƷID")) & "|" & str��������
                End If
            Next
        Next
    End With

    gstrSql = "Zl_��ѡҩƷ����_Update('" & str�������� & "')"
    
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Call vsfList_EnterCell
    Call EditColor
    
    MsgBox "����ɹ���", vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Call IniGrid
    Call FillVSF
    Call ShowData
    Call EditColor
End Sub

Private Sub IniGrid()
    With vsfList
        .Editable = flexEDNone
        .Rows = 1
        .ColWidth(0) = 350
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = 400
        .AllowSelection = False '���ܶ�ѡ
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
    End With

    With vsf����
        .Editable = flexEDNone
        .Rows = 2
        .ColWidth(0) = 350
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = 400
        .AllowSelection = False '���ܶ�ѡ
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .Cell(flexcpBackColor, 1, .ColIndex("����ҩƷ"), 1, .ColIndex("����ҩƷ")) = mcstEditColor
    End With
End Sub

Private Sub FillVSF()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSql = "Select a.ҩƷid, b.����, b.����, b.���, b.���� As ������, c.���� As ��Ӧ��,n.���� As ��Ʒ��" & vbNewLine & _
                    "From ҩƷ��� A, �շ���ĿĿ¼ B, ��Ӧ�� C, �շ���Ŀ���� N" & vbNewLine & _
                    "Where a.ҩƷid = b.Id And a.�ϴι�Ӧ��id = c.Id(+) And b.Id = n.�շ�ϸĿid(+)" & vbNewLine & _
                    "      And n.����(+) = 1 And n.����(+) = 3 And a.�Ƿ�����ɹ� = 1" & vbNewLine & _
                    "Order By b.����"


    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "������ɹ�ҩƷ")
    
    With vsfList
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("���")) = .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("ҩƷID")) = rsTemp!ҩƷID
            .TextMatrix(.Rows - 1, .ColIndex("��ѡҩƷ")) = "[" & rsTemp!���� & "]" & rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("��Ʒ��")) = NVL(rsTemp!��Ʒ��)
            .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTemp!���
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = NVL(rsTemp!��Ӧ��)
            
            rsTemp.MoveNext
        Loop
        
    End With
    
    Call VsfRowHeight(vsfList)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowData()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vsf���ݻ���.Rows = 1
    
    For i = 1 To vsfList.Rows - 1

            gstrSql = "Select b.����, b.����, b.���, b.���� As ������, c.���� As ��Ӧ��, d.���, d.��ѡҩƷid, d.����ҩƷid, n.���� As ��Ʒ��" & vbNewLine & _
                            "From ҩƷ��� A, �շ���ĿĿ¼ B, ��Ӧ�� C, ��ѡҩƷ���� D, �շ���Ŀ���� N" & vbNewLine & _
                            "Where a.ҩƷid = b.Id And a.�ϴι�Ӧ��id = c.Id(+) And b.Id = d.����ҩƷid" & vbNewLine & _
                            "      And b.Id = n.�շ�ϸĿid(+) And n.����(+) = 1 And n.����(+) = 3 And d.��ѡҩƷid = [1]" & vbNewLine & _
                            "Order By d.���"

            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "����ҩƷ", Val(vsfList.TextMatrix(i, vsfList.ColIndex("ҩƷID"))))
            
            With vsf���ݻ���
        
                Do While Not rsTemp.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTemp!���
                    .TextMatrix(.Rows - 1, .ColIndex("��ѡҩƷID")) = rsTemp!��ѡҩƷID
                    .TextMatrix(.Rows - 1, .ColIndex("����ҩƷID")) = rsTemp!����ҩƷID
                    .TextMatrix(.Rows - 1, .ColIndex("����ҩƷ")) = "[" & rsTemp!���� & "]" & rsTemp!����
                    .TextMatrix(.Rows - 1, .ColIndex("��Ʒ��")) = NVL(rsTemp!��Ʒ��)
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTemp!���
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
                    .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = NVL(rsTemp!��Ӧ��)
                    
                    rsTemp.MoveNext
                Loop
        
            End With
        Next
        
        If vsfList.Rows > 1 Then vsfList.Row = 1

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ҩƷID = 0
    mlng��ѡҩƷID = 0
    mlngID = 0
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub vsf����_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsf����
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("��Ӧ��") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("��Ӧ��") Then
                .Row = .Row + 1
                .Col = .ColIndex("����ҩƷ")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("����ҩƷ")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete
        End If
    End With
End Sub

Private Sub vsf����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsf����.ColIndex("����ҩƷ") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsf����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsf����.hwnd) '��ȡλ��
    dblLeft = vRect.Left + vsf����.CellLeft
    dblTop = vRect.Top + vsf����.CellTop + vsf����.CellHeight + 3300
    With vsf����
        mlng����ҩƷID = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Col = .ColIndex("����ҩƷ") And .EditText = "" Then Exit Sub
        If Col = .ColIndex("����ҩƷ") And InStr(1, .EditText, "[") = 0 Then
        gstrSql = "Select Distinct i.Id, i.����, i.����, i.���, i.���� As ������, c.���� As ��Ӧ��, m.��Ʒ��" & vbNewLine & _
                        "From �շ���ĿĿ¼ I, �շ���Ŀ���� N, ҩƷ��� A, ��Ӧ�� C, (Select �շ�ϸĿid, ���� As ��Ʒ�� From �շ���Ŀ���� Where ���� = 1 And ���� = 3) M" & vbNewLine & _
                        "Where i.Id = n.�շ�ϸĿid And i.Id = a.ҩƷid And a.�ϴι�Ӧ��id = c.Id(+) And i.Id = m.�շ�ϸĿid(+) And i.��� In ('5', '6') And" & vbNewLine & _
                        "      (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "      (i.���� Like [1] Or n.���� Like [2] Or n.���� Like [2]) And Nvl(a.�Ƿ�����ɹ�, 0) = 0" & vbNewLine & _
                        "Order By i.����"

            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "����ҩƷ", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%", gstrMatch & UCase(.EditText) & "%")
            
            If blnCancel = True Then
                Exit Sub
            End If
  
            If rsRecord Is Nothing Then
                MsgBox "û���ҵ��ö���ҩƷ��", vbInformation, gstrSysName
                Exit Sub
            Else
                mlngID = rsRecord!ID
                If CheckDub = False Then
                    .EditText = "[" & rsRecord!���� & "]" & rsRecord!����
                    .TextMatrix(.Row, .ColIndex("���")) = .Row
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("����ҩƷ")) = "[" & rsRecord!���� & "]" & rsRecord!����
                    .TextMatrix(.Row, .ColIndex("��Ʒ��")) = NVL(rsRecord!��Ʒ��)
                    .TextMatrix(.Row, .ColIndex("���")) = rsRecord!���
                    .TextMatrix(.Row, .ColIndex("������")) = NVL(rsRecord!������)
                    .TextMatrix(.Row, .ColIndex("��Ӧ��")) = NVL(rsRecord!��Ӧ��)
                    
                    Call UpDate
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Cell(flexcpBackColor, .Row, .ColIndex("����ҩƷ"), .Rows - 1, .ColIndex("����ҩƷ")) = mcstEditColor
                    Call VsfRowHeight(vsf����)
                Else
                    MsgBox "�Ѿ��и�ҩƷ��", vbInformation, gstrSysName
                End If
            End If
            
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf����_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsf����
        If .Col = .ColIndex("����ҩƷ") Then
            .ColComboList(.ColIndex("����ҩƷ")) = "|..."
        Else
            .ColComboList(.ColIndex("����ҩƷ")) = ""
        End If
    End With
End Sub

Private Sub vsf����_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsf����
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsf����_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsf����
        .EditMaxLength = 50
    End With
End Sub


Private Sub vsf����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsf����.hwnd) '��ȡλ��
    dblLeft = vRect.Left + vsf����.CellLeft
    dblTop = vRect.Top + vsf����.CellTop + vsf����.CellHeight + 3300
    With vsf����
        mlng����ҩƷID = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
        If Col = .ColIndex("����ҩƷ") Then
            gstrSql = "Select i.Id, i.����, i.����, i.���, i.���� As ������, c.���� As ��Ӧ��, n.���� As ��Ʒ��" & vbNewLine & _
                            "From �շ���ĿĿ¼ I, ҩƷ��� A, ��Ӧ�� C, �շ���Ŀ���� N" & vbNewLine & _
                            "Where i.Id = a.ҩƷid And a.�ϴι�Ӧ��id = c.Id(+) And i.��� In ('5', '6') And i.Id = n.�շ�ϸĿid(+)" & vbNewLine & _
                            "      And n.����(+) = 1 And n.����(+) = 3 And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                            "      And Nvl(a.�Ƿ�����ɹ�, 0) = 0 Order By i.����"

            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "����ҩƷ", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)

            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngID = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(.Row, .ColIndex("���")) = .Row
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("����ҩƷ")) = "[" & rsRecord!���� & "]" & rsRecord!����
                    .TextMatrix(.Row, .ColIndex("��Ʒ��")) = NVL(rsRecord!��Ʒ��)
                    .TextMatrix(.Row, .ColIndex("���")) = rsRecord!���
                    .TextMatrix(.Row, .ColIndex("������")) = NVL(rsRecord!������)
                    .TextMatrix(.Row, .ColIndex("��Ӧ��")) = NVL(rsRecord!��Ӧ��)
                    
                    Call UpDate
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Cell(flexcpBackColor, .Row, .ColIndex("����ҩƷ"), .Rows - 1, .ColIndex("����ҩƷ")) = mcstEditColor
                    Call VsfRowHeight(vsf����)
                Else
                    MsgBox "�Ѿ��и�ҩƷ��", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
    
End Sub

Private Sub vsf����_EnterCell()
    
    With vsf����
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstEditColor Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
        mlng����ҩƷID = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    
    With vsfList
        If Val(.TextMatrix(.Row, .ColIndex("ҩƷID"))) = 0 Then Exit Sub
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            .CellBorderRange .Row, 0, .Row, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
        mlng��ѡҩƷID = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
        Call ShowGrid(Val(.TextMatrix(.Row, .ColIndex("ҩƷID"))))
    End With
    
End Sub

Private Function CheckDub() As Boolean
    '����Ƿ���ڸ��б굥λ���Ƿ����ҩƷ
    Dim i As Integer

    With vsf���ݻ���
        For i = 1 To .Rows - 1
            If Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))) = Val(.TextMatrix(i, .ColIndex("��ѡҩƷID"))) And _
                .TextMatrix(i, .ColIndex("����ҩƷid")) = mlngID Then
                CheckDub = True
                Exit Function
            End If
        Next
    End With
    CheckDub = False

End Function

Private Sub Delete()
    Dim i As Integer
    
    With vsf���ݻ���
        For i = 1 To .Rows - 1
            If Val(vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("ҩƷID"))) = Val(.TextMatrix(i, .ColIndex("����ҩƷID"))) And _
                Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))) = Val(.TextMatrix(i, .ColIndex("��ѡҩƷID"))) Then
                .RemoveItem i
                Exit For
            End If
        Next
    End With
    
    With vsf����
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ҩƷID"))) = 0 Then Exit Sub
        
        If .Rows - 1 = 1 Then
            For i = 1 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
        Else
            .RemoveItem .Row
            Call vsf_ResetSerial
        End If
    End With
    
End Sub

Private Sub vsf_ResetSerial()
    Dim i As Integer
    With vsf����
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("���")) = i
        Next
    End With
End Sub

Private Sub UpDate()
    Dim i As Integer
    
    With vsf���ݻ���
        If vsf����.Row = vsf����.Rows - 1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("��ѡҩƷID")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))
            .TextMatrix(.Rows - 1, .ColIndex("����ҩƷID")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("ҩƷID"))
            .TextMatrix(.Rows - 1, .ColIndex("����ҩƷ")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("����ҩƷ"))
            .TextMatrix(.Rows - 1, .ColIndex("��Ʒ��")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("��Ʒ��"))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("���"))
            .TextMatrix(.Rows - 1, .ColIndex("������")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("������"))
            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = vsf����.TextMatrix(vsf����.Rows - 1, vsf����.ColIndex("��Ӧ��"))
        Else
            For i = 1 To .Rows - 1
                If mlng����ҩƷID = Val(.TextMatrix(i, .ColIndex("����ҩƷID"))) And _
                    mlng��ѡҩƷID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))) Then
                        .TextMatrix(i, .ColIndex("��ѡҩƷID")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))
                        .TextMatrix(i, .ColIndex("����ҩƷID")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("ҩƷID"))
                        .TextMatrix(i, .ColIndex("����ҩƷ")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("����ҩƷ"))
                        .TextMatrix(i, .ColIndex("��Ʒ��")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��Ʒ��"))
                        .TextMatrix(i, .ColIndex("���")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("���"))
                        .TextMatrix(i, .ColIndex("������")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("������"))
                        .TextMatrix(i, .ColIndex("��Ӧ��")) = vsf����.TextMatrix(vsf����.Row, vsf����.ColIndex("��Ӧ��"))
                    Exit For
                End If
            Next
        End If
    End With
End Sub


Private Sub ShowGrid(ByVal lngҩƷID As Long)
    Dim n As Long
    
    With vsf����
        .Rows = 1
        .Rows = 2
        .Cell(flexcpBackColor, 1, .ColIndex("����ҩƷ"), 1, .ColIndex("����ҩƷ")) = mcstEditColor
        For n = 1 To vsf���ݻ���.Rows - 1
            If lngҩƷID = Val(vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("��ѡҩƷID"))) Then
                .TextMatrix(.Rows - 1, .ColIndex("���")) = .Rows - 1
                .TextMatrix(.Rows - 1, .ColIndex("ҩƷID")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("����ҩƷID"))
                .TextMatrix(.Rows - 1, .ColIndex("����ҩƷ")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("����ҩƷ"))
                .TextMatrix(.Rows - 1, .ColIndex("��Ʒ��")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("��Ʒ��"))
                .TextMatrix(.Rows - 1, .ColIndex("���")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("���"))
                .TextMatrix(.Rows - 1, .ColIndex("������")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("������"))
                .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("��Ӧ��"))
                
                .Rows = .Rows + 1
                .Cell(flexcpBackColor, 1, .ColIndex("����ҩƷ"), .Rows - 1, .ColIndex("����ҩƷ")) = mcstEditColor
            End If
        Next
    End With
    Call VsfRowHeight(vsf����)
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long

    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput

        gstrSql = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B,ҩƷ��� C " & _
                  "Where a.id=c.ҩƷID And A.Id =B.�շ�ϸĿid And A.��� In ('5','6') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) and c.�Ƿ�����ɹ�=1 " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSql, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����

        For lngRow = 1 To vsfList.Rows - 1
            lngFindRow = vsfList.FindRow(strҩ��, lngRow, CLng(vsfList.ColIndex("��ѡҩƷ")), True, True)
            If lngFindRow > 0 Then
                vsfList.Row = lngFindRow
                vsfList.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '��ѯ�����ݺ���ƶ�����һ�����˳����β�ѯ
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub EditColor()
    Dim i As Long, n As Long
    Dim bln�Ƿ����ö��� As Boolean
    
    With vsfList
        For i = 1 To .Rows - 1
            bln�Ƿ����ö��� = False
            For n = 1 To vsf���ݻ���.Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ҩƷID"))) = Val(vsf���ݻ���.TextMatrix(n, vsf���ݻ���.ColIndex("��ѡҩƷID"))) Then
                    bln�Ƿ����ö��� = True
                    Exit For
                End If
            Next
            If bln�Ƿ����ö��� Then
                .Cell(flexcpForeColor, i, .ColIndex("���"), i, .ColIndex("��Ӧ��")) = mcstBachColor
            Else
                .Cell(flexcpForeColor, i, .ColIndex("���"), i, .ColIndex("��Ӧ��")) = mcstNotBachColor
            End If
        Next
    End With
End Sub

Private Sub VsfRowHeight(ByVal VsfObj As VSFlexGrid)
    Dim i As Long
    With VsfObj
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub
