VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmParAdviceSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "·��ѡ��"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295
   Icon            =   "frmParAdviceSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7080
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8295
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����·����Ŀ����˳��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1095
         TabIndex        =   8
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    ·����Ŀ����ҽ��ʱȱʡ��·�����иý׶ζ���ķ��༰��Ŀ˳���г��������Ȱ��±�������˳�����У�ÿ������ҽ��ʱҲ���Ե���˳��"
         Height          =   360
         Left            =   1095
         TabIndex        =   7
         Top             =   360
         Width           =   6165
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmParAdviceSort.frx":038A
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   7680
      Picture         =   "frmParAdviceSort.frx":6504
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   7080
      Picture         =   "frmParAdviceSort.frx":69B5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.PictureBox picAddRow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   220
      Left            =   6000
      Picture         =   "frmParAdviceSort.frx":6E6E
      ScaleHeight     =   225
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   360
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   6435
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6900
      _cx             =   12171
      _cy             =   11351
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmParAdviceSort.frx":71F8
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmParAdviceSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mbytFun  As Byte     '0=�ٴ�·��ģ�����,1=ҽ��վ����
Private Enum CNAME
    c˳�� = 0
    c��Ч = 1
    c������� = 2
    c�������� = 3
    c��ҩ���� = 4
    c���� = 5
End Enum


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With vsItem
            If Index = 0 And .Row > .FixedRows Then
                .RowPosition(.Row) = .Row - 1
                .TextMatrix(.Row, c˳��) = .TextMatrix(.Row, c˳��) + 1
                .TextMatrix(.Row - 1, c˳��) = .TextMatrix(.Row - 1, c˳��) - 1
                .Row = .Row - 1
            ElseIf Index = 1 And .Row < .Rows - 1 Then
                .RowPosition(.Row) = .Row + 1
                .TextMatrix(.Row, c˳��) = .TextMatrix(.Row, c˳��) - 1
                .TextMatrix(.Row + 1, c˳��) = .TextMatrix(.Row + 1, c˳��) + 1
                .Row = .Row + 1
            End If
    End With
End Sub

Private Sub cmdOK_Click()
    If Not (vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c��Ч) = "" And vsItem.TextMatrix(1, CNAME.c�������) = "" _
            And vsItem.TextMatrix(1, CNAME.c��������) = "" And vsItem.TextMatrix(1, CNAME.c��ҩ����) = "") Then
        If CheckData = False Then Exit Sub
    End If
    
    Call SaveData
    Unload Me
End Sub

Private Function CheckData() As Boolean
    Dim i As Long, str�������� As String, str��ҩ���� As String
    Dim rsSQL As ADODB.Recordset, strKey As String
    
    
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "�к�", adBigInt
    rsSQL.Fields.Append "ֵ", adVarChar, 200, adFldIsNullable
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    With vsItem
        .Cell(flexcpBackColor, .FixedRows, c˳��, .Rows - 1, .Cols - 1) = vbWhite
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c��Ч) = "" Then
                MsgBox "��ѡ��ҽ����Ч��", vbInformation, gstrSysName
                .Select i, c��Ч
                Exit Function
            ElseIf .TextMatrix(i, c�������) = "" Then
                MsgBox "��ѡ��������Ŀ���", vbInformation, gstrSysName
                .Select i, c�������
                Exit Function
            ElseIf .TextMatrix(i, c��������) = "" Then
                If .Cell(flexcpData, i, c�������) = "H" Or .Cell(flexcpData, i, c�������) = "E" Then
                    MsgBox "��ѡ��������͡�", vbInformation, gstrSysName
                    .Select i, c��������
                    Exit Function
                End If
            ElseIf .TextMatrix(i, c��ҩ����) = "" Then
                If .TextMatrix(i, c�������) = "��ҩ�г�ҩ" Then
                    MsgBox "��ѡ���ҩ���ࡣ", vbInformation, gstrSysName
                    .Select i, c��ҩ����
                    Exit Function
                End If
            End If
        
    
            '����ظ�ֵ
            If .TextMatrix(i, c��������) = "" Then
                str�������� = "Null"
            Else
                str�������� = .Cell(flexcpData, i, c��������)
            End If
            
            If .TextMatrix(i, c��ҩ����) = "" Then
                str��ҩ���� = "Null"
            Else
                str��ҩ���� = .Cell(flexcpData, i, c��ҩ����)
            End If
            strKey = .Cell(flexcpData, i, c��Ч) & "," & .Cell(flexcpData, i, c�������) & "," & str�������� & "," & str��ҩ����
            
            rsSQL.Filter = "ֵ='" & strKey & "'"
            If rsSQL.RecordCount > 0 Then
                MsgBox "��" & i & "�����" & rsSQL!�к� & "�е������ظ���", vbInformation, gstrSysName
                .Cell(flexcpBackColor, Val(rsSQL!�к�), c˳��, Val(rsSQL!�к�), .Cols - 1) = &H80C0FF
                .Select i, c��Ч
                Exit Function
            Else
                rsSQL.AddNew
                rsSQL!�к� = i
                rsSQL!ֵ = strKey
                rsSQL.Update
            End If
        Next
    End With
    CheckData = True
End Function

Private Sub SaveData()
    Dim strSQL As String
    Dim i As Long, str�������� As String, str��ҩ���� As String
    Dim colSQL As New Collection, blnTrans As Boolean, blnSetup As Boolean
    Dim intOnlyDel As Integer
    Dim strTmp As String
    
    On Error GoTo errH
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c��������) = "" Then
                str�������� = "Null"
            Else
                str�������� = "'" & .Cell(flexcpData, i, c��������) & "'"
            End If
            
            If .TextMatrix(i, c��ҩ����) = "" Then
                str��ҩ���� = "Null"
            Else
                str��ҩ���� = .Cell(flexcpData, i, c��ҩ����)
            End If

            If vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c��Ч) = "" And vsItem.TextMatrix(1, CNAME.c�������) = "" _
                    And vsItem.TextMatrix(1, CNAME.c��������) = "" And vsItem.TextMatrix(1, CNAME.c��ҩ����) = "" Then
                intOnlyDel = 1
            Else
                intOnlyDel = 0
            End If
            strSQL = "Zl_·����Ŀ˳��_Insert(" & .TextMatrix(i, c˳��) & "," & _
                IIF(.Cell(flexcpData, i, c��Ч) = "", "null", .Cell(flexcpData, i, c��Ч)) & _
                ",'" & .Cell(flexcpData, i, c�������) & "'," & str�������� & "," & str��ҩ���� & "," & _
                intOnlyDel & ")"
            colSQL.Add strSQL, "C" & colSQL.Count + 1
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.Count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim blnParSet As Boolean
    Dim lngDays As Long
    Dim strTmp As String
    
    If mbytFun = 0 Then
        Me.Caption = "·��ѡ��"
        lblInfo.Caption = "ҽ����·����Ŀ����˳��"
        lblNote.Caption = "    ·����Ŀ����ҽ��ʱȱʡ��·�����иý׶ζ���ķ��༰��Ŀ˳���г��������Ȱ��±�������˳�����У�ÿ������ҽ��ʱҲ���Ե���˳��"
    Else
        Me.Caption = "ҽ����������"
        lblInfo.Caption = "ҽ���´���Զ�����"
        lblNote.Caption = "    ҽ������ǰ���Ա����¿���ҽ�����Զ����±�������˳�����У������Ҳ����ʹ��ҽ��˳�����������������˳��"
    End If
    
    picAddRow.Visible = False
    Call InitData
    Call LoadData
    
    '�������Ƶ���һ��
    If vsItem.Rows > 0 Then vsItem.Row = 1
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select a.˳��,a.ҽ����Ч,a.������� as ������,a.ִ�з���,a.��������,b.���� as ������� From ·����Ŀ˳�� a,������Ŀ��� b " & _
        "Where a.������� = b.���� Order by a.˳��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "·����Ŀ˳��")
    
    If rsTmp.RecordCount = 0 Then vsItem.TextMatrix(1, c˳��) = 1: Exit Sub
    
    With vsItem
        .redraw = False
        .Rows = .FixedRows + rsTmp.RecordCount
        i = .FixedRows
        While Not rsTmp.EOF
            .TextMatrix(i, c˳��) = i
            .TextMatrix(i, c��Ч) = IIF(rsTmp!ҽ����Ч = 0, "����", "����")
            .Cell(flexcpData, i, c��Ч) = Val(rsTmp!ҽ����Ч)
            
            .Cell(flexcpData, i, c�������) = CStr(rsTmp!������)  'ҩƷ���Ǵ�Ϊ������ı���E
            
            If rsTmp!������ = "E" And Val("" & rsTmp!��������) = 2 Then
                .TextMatrix(i, c�������) = "��ҩ�г�ҩ"
                
            ElseIf rsTmp!������ = "E" And Val("" & rsTmp!��������) = 4 Then
                .TextMatrix(i, c�������) = "�в�ҩ"
                
            Else
                .TextMatrix(i, c�������) = rsTmp!�������
            End If
            
            '��ֻ֧�֣������ࣺ0-��ͨ;1-��������;2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;5-��������;6-�ɼ�����(����);7-��Ѫ����(Ѫ��);8-��Ѫ;����
            '            �����ࣺ0-�����棻1-����ȼ���
            If Not IsNull(rsTmp!��������) And (rsTmp!������ = "H" Or rsTmp!������ = "E") Then
                If rsTmp!������ = "H" Then
                    .TextMatrix(i, c��������) = IIF(rsTmp!�������� = 0, "������", "����ȼ�")
                Else
                     .TextMatrix(i, c��������) = Choose(Val(rsTmp!��������) + 1, "��ͨ", "��������", "��ҩ����", "��ҩ�巨", "��ҩ�÷�", "��������", "�ɼ�����", "��Ѫ����", "��Ѫ;��")
                End If
                .Cell(flexcpData, i, c��������) = Val(rsTmp!��������)
            End If
            
            If Not IsNull(rsTmp!ִ�з���) Then
                .TextMatrix(i, c��ҩ����) = Choose(rsTmp!ִ�з��� + 1, "����", "��Һ", "ע��", "Ƥ��", "�ڷ�")
                .Cell(flexcpData, i, c��ҩ����) = Val("" & rsTmp!ִ�з���)
            End If
            
            If rsTmp!������ = "Z" And Val("" & rsTmp!��������) = 4 Then .TextMatrix(i, c��������) = "����": .Cell(flexcpData, i, c��������) = Val(rsTmp!��������)
            i = i + 1
            rsTmp.MoveNext
        Wend
        .redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset, strTmp As String
    
    Set rsTmp = Get�������
    strTmp = "#E;��ҩ�г�ҩ|#E;�в�ҩ"  '�̶����ƣ���������洢
    While Not rsTmp.EOF
        strTmp = strTmp & "|#" & rsTmp!���� & ";" & rsTmp!����
        rsTmp.MoveNext
    Wend
    
    With vsItem
        
        .ColComboList(c��Ч) = "#1;����|#0;����"
        .ColComboList(c�������) = strTmp
        .ColComboList(c��ҩ����) = "#0;����|#1;��Һ|#2;ע��|#3;Ƥ��|#4;�ڷ�"
        .Rows = .FixedRows
        .Rows = .FixedRows + 1 '��ʼһ����
    End With
End Sub

Private Function Get�������() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not In('5','6','7')"
    Set Get������� = zlDatabase.OpenSQLRecord(strSQL, "�������")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picAddRow_Click()
    Dim i As Long
    
    If vsItem.Row = vsItem.Rows - 1 Then
        vsItem.Rows = vsItem.Rows + 1
        vsItem.TextMatrix(vsItem.Rows - 1, c˳��) = vsItem.Rows - 1
        vsItem.Select vsItem.Rows - 1, c��Ч
    Else
        i = vsItem.Row
        vsItem.AddItem "", i
        Call Reset���
        vsItem.Select i, c��Ч
    End If
    
End Sub

Private Sub Reset���()
    Dim i As Long
    
    For i = vsItem.FixedRows To vsItem.Rows - 1
        vsItem.TextMatrix(i, c˳��) = i
    Next
End Sub


Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsItem.ComboData = "" Then   'δѡ��ʱ�뿪����
        vsItem.TextMatrix(Row, Col) = CStr(vsItem.Tag)
        Exit Sub
    End If
    
    With vsItem
        If .Tag <> "" Then
            If .Tag = CStr(.ComboItem) Then Exit Sub
        End If
        .TextMatrix(Row, Col) = .ComboItem
        .Cell(flexcpData, Row, Col) = .ComboData
        
        If Col = c������� Then
            If .TextMatrix(Row, c�������) = "��ҩ�г�ҩ" Then
                .TextMatrix(Row, c��������) = "��ҩ����"
                .Cell(flexcpData, Row, c��������) = 2
            
            ElseIf .TextMatrix(Row, c�������) = "�в�ҩ" Then
                .TextMatrix(Row, c��������) = "��ҩ�÷�"
                .Cell(flexcpData, Row, c��������) = 4
                
            Else
                .TextMatrix(Row, c��������) = ""
                .Cell(flexcpData, Row, c��������) = ""
            End If
            
            .TextMatrix(Row, c��ҩ����) = ""
            .Cell(flexcpData, Row, c��ҩ����) = ""
            
        ElseIf Col = c�������� Then
            .TextMatrix(Row, c��ҩ����) = ""
            .Cell(flexcpData, Row, c��ҩ����) = ""
        ElseIf Col = c��Ч Then
            If .TextMatrix(Row, c�������) = "����" And .TextMatrix(Row, c��Ч) = "����" Then
                .TextMatrix(Row, c��������) = ""
                .Cell(flexcpData, Row, c��������) = ""
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (OldRow <> NewRow Or OldRow = NewRow And OldRow = 1) And NewRow > vsItem.FixedRows - 1 Then
        If Me.Visible Then
            If picAddRow.Visible = False Then picAddRow.Visible = True
        End If
        picAddRow.Top = vsItem.Top + vsItem.Cell(flexcpTop, NewRow, c����) + 30
        picAddRow.Left = vsItem.Left + vsItem.Cell(flexcpLeft, NewRow, c����) + 120
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '��ֻ֧�֣������ࣺ0-��ͨ;1-��������;2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;5-��������;6-�ɼ�����(����);7-��Ѫ����(Ѫ��);8-��Ѫ;����
            '            �����ࣺ0-�����棻1-����ȼ���
    With vsItem
        .Tag = .TextMatrix(Row, Col)  '����AfterEdit���ж��Ƿ�ı���ֵ
        If Col = c���� Then
            Cancel = True
        ElseIf Col = c�������� Then '���ƺͻ��������
            If .Cell(flexcpData, Row, c�������) = "H" Then
                .ComboList = "#0;������|#1;����ȼ�"
                
            ElseIf .TextMatrix(Row, c�������) = "��ҩ�г�ҩ" Or .TextMatrix(Row, c�������) = "�в�ҩ" Then
                .ComboList = ""
                Cancel = True
            ElseIf .Cell(flexcpData, Row, c�������) = "E" Then
                .ComboList = "#0;��ͨ|#1;��������|#5;��������"
            ElseIf .Cell(flexcpData, Row, c�������) = "Z" And .TextMatrix(Row, c��Ч) = "����" Then
                .ComboList = "#0;|#4;����"
            Else
                .ComboList = ""
                Cancel = True
            End If
        ElseIf Col = c��ҩ���� Then 'ҩƷ������
            If Not (.TextMatrix(Row, c�������) = "��ҩ�г�ҩ") Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsItem_ChangeEdit()
    Call vsItem_AfterEdit(vsItem.Row, vsItem.Col)
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        If vsItem.Row = 0 Then Exit Sub
        If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        With vsItem
        
            If .Rows > 2 Then
                vsItem.RemoveItem vsItem.Row
                Call Reset���
            ElseIf .Rows = 2 Then
                If .TextMatrix(1, CNAME.c��Ч) = "" And .TextMatrix(1, CNAME.c�������) = "" _
                        And .TextMatrix(1, CNAME.c��������) = "" And .TextMatrix(1, CNAME.c��ҩ����) = "" Then
                    MsgBox "û�п�ɾ�������ˡ�", vbInformation, gstrSysName
                Else
                    .TextMatrix(1, CNAME.c��Ч) = ""
                    .TextMatrix(1, CNAME.c�������) = ""
                    .TextMatrix(1, CNAME.c��������) = ""
                    .TextMatrix(1, CNAME.c��ҩ����) = ""
                End If
            End If
        
        End With
       
    End If
End Sub

Private Sub EnterNextCell()
   
    With vsItem
        If .Col = .Cols - 1 And .Row < .Rows - 1 Then
            .Select .Row + 1, c��Ч
        ElseIf .Col < .Cols - 1 Then
            .Col = .Col + 1
        End If
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsItem.ComboIndex <> -1 Then
            Call vsItem_KeyPress(13)
        End If
    End If
End Sub

