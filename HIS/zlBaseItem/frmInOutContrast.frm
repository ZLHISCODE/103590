VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInOutContrast 
   Caption         =   "�����������"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   Icon            =   "frmInOutContrast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10605
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   75
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   10635
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8040
      TabIndex        =   5
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   9300
      TabIndex        =   4
      Top             =   5070
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNo 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5535
      _cx             =   9763
      _cy             =   7858
      Appearance      =   1
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInOutContrast.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfClass 
      Height          =   4455
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   4695
      _cx             =   8281
      _cy             =   7858
      Appearance      =   1
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInOutContrast.frx":68CC
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
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl���ݷ��� 
      AutoSize        =   -1  'True
      Caption         =   "���ݷ���"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmInOutContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsClass As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTemp As String
    Dim intRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim lng�� As Long
    Dim lng�� As Long
    
    On Error GoTo errHandle
    
    With vsfClass
        If Trim(.TextMatrix(1, .ColIndex("�������"))) = "" Or Trim(.TextMatrix(1, .ColIndex("�������"))) = "" Then
            MsgBox "��������һ�����������ܱ��棡", vbInformation, gstrSysName
            Exit Sub
        End If
        '����Ƿ����ظ���
        For i = 1 To .Rows - 1
            For j = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�������")) = .TextMatrix(j, .ColIndex("�������")) And i <> j Then
                    MsgBox "һ�������ֻ�ܶ�Ӧһ�������", vbInformation, gstrSysName
                    .Row = j
                    .Col = .ColIndex("�������")
                    Exit Sub
                End If
                
                If .TextMatrix(i, .ColIndex("�������")) = .TextMatrix(j, .ColIndex("�������")) And i <> j Then
                    MsgBox "һ�������ֻ�ܶ�Ӧһ�������", vbInformation, gstrSysName
                    .Row = j
                    .Col = .ColIndex("�������")
                    Exit Sub
                End If
            Next
        Next
        
        strTemp = vsfNo.TextMatrix(vsfNo.Row, vsfNo.ColIndex("����id")) & "|"
        
        For intRow = 1 To .Rows - 1
            If Trim(.TextMatrix(intRow, .ColIndex("�������"))) <> "" And Trim(.TextMatrix(intRow, .ColIndex("�������"))) <> "" Then
                lng�� = Mid(.TextMatrix(intRow, .ColIndex("�������")), 1, InStr(1, .TextMatrix(intRow, .ColIndex("�������")), "-") - 1)
                lng�� = Mid(.TextMatrix(intRow, .ColIndex("�������")), 1, InStr(1, .TextMatrix(intRow, .ColIndex("�������")), "-") - 1)
                strTemp = strTemp & lng�� & "," & lng�� & ";"
            End If
        Next
    End With
    
    gstrSQL = "Zl_���������_Insert('" & strTemp & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "���������"
    MsgBox "����ɹ���", vbInformation, gstrSysName
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call GetList
    Call InitVsfGridFlex
End Sub

Public Sub ShowMe(ByVal fraPar As Form)
    Me.Show vbModal, fraPar
End Sub

Private Sub GetList()
    '��ȡ����б�����
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ���� As ����, ����, ˵�� From ҩƷ���ݷ��� Where ���� = 6"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ���ݷ���")
    
    With vsfNo
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("����id")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("˵��")) = rsTemp!˵��
            
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub vsfClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With vsfClass
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .Rows - 1 Then
                    If .TextMatrix(.Row, 0) <> "" And .TextMatrix(.Row, 1) <> "" Then
                        .Row = .Row + 1
                        .Col = 0
                    End If
                Else
                    If .TextMatrix(.Row, 0) <> "" And .TextMatrix(.Row, 1) <> "" Then
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = 0
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeyDelete Then
        If Trim(vsfClass.TextMatrix(1, vsfClass.ColIndex("�������"))) = "" And Trim(vsfClass.TextMatrix(1, vsfClass.ColIndex("�������"))) = "" Then Exit Sub
        If MsgBox("��ɾ����ǰѡ���У��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        vsfClass.RemoveItem vsfClass.Row
        If vsfClass.Rows <= 1 Then vsfClass.Rows = vsfClass.Rows + 1
    End If
End Sub

Private Sub vsfNo_EnterCell()
    Dim lngNo As Long
    Dim intRow As Integer
    Dim lng�����id As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    With vsfNo
        vsfClass.Rows = 1
        
        If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("����id")) <> "" Then
            lngNo = .TextMatrix(.Row, .ColIndex("����id"))
            '�ȼ����������� �����
            gstrSQL = "Select b.Id || '-' || b.���� As ���� From ��������� A, ҩƷ������ B Where a.�����id = b.Id And a.���� = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�����", lngNo)
            Do While Not rsTemp.EOF
                vsfClass.Rows = vsfClass.Rows + 1
                vsfClass.TextMatrix(vsfClass.Rows - 1, vsfClass.ColIndex("�������")) = rsTemp!����
                rsTemp.MoveNext
            Loop
            
            '�ȼ����������� �����
            For intRow = 1 To vsfClass.Rows - 1
                If vsfClass.TextMatrix(intRow, vsfClass.ColIndex("�������")) <> "" Then
                    lng�����id = Mid(vsfClass.TextMatrix(intRow, vsfClass.ColIndex("�������")), 1, InStr(1, vsfClass.TextMatrix(intRow, vsfClass.ColIndex("�������")), "-") - 1)
                    
                    gstrSQL = "Select b.id || '-' || b.���� as ���� From ��������� A, ҩƷ������ b Where a.�����id=b.id and a.�����id=[2] and a.����=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�����", lngNo, lng�����id)
                    vsfClass.TextMatrix(intRow, vsfClass.ColIndex("�������")) = rsTemp!����
                End If
            Next
            
            '�ٶ��а�����
            gstrSQL = "Select a.���id || '-' || b.���� as ����, b.ϵ��" & vbNewLine & _
                    " From ҩƷ�������� A, ҩƷ������ B" & vbNewLine & _
                    " Where a.���id = b.Id And a.���� = [1]" & vbNewLine & _
                    " Order By b.ϵ��"
            Set mrsClass = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ������", lngNo)
                        
            mrsClass.Filter = " ϵ��=1"
            vsfClass.ColComboList(vsfClass.ColIndex("�������")) = vsfClass.BuildComboList(mrsClass, "����")
            mrsClass.Filter = " ϵ��=-1"
            vsfClass.ColComboList(vsfClass.ColIndex("�������")) = vsfClass.BuildComboList(mrsClass, "����")
            If vsfClass.Rows = 1 Then
                vsfClass.Rows = 2
            End If
        Else
            vsfClass.Rows = 1
        End If
    End With

    Call InitVsfGridFlex
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsfGridFlex()

    With vsfNo
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ������и�
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .AllowSelection = False '���ܶ�ѡ
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExtendLastCol = True '���һ�������
        .WordWrap = True
        .AutoSize .ColIndex("˵��"), , False, 0 = True
        .ScrollBars = flexScrollBarVertical '�����������ȡ����
    End With
    
    With vsfClass
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .AllowSelection = False '���ܶ�ѡ
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
    End With
End Sub
