VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIdxInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "�������ȱʧ���ҺͲ���"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12600
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8130
      Left            =   5505
      ScaleHeight     =   8130
      ScaleWidth      =   7095
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdExecute 
         Caption         =   "ִ��(&E)"
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   7680
         Width           =   1095
      End
      Begin VB.OptionButton optFKey 
         Caption         =   "�������"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtToolInfo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmIdxInfo.frx":0000
         Top             =   120
         Width           =   6615
      End
      Begin VB.Frame fraIdx 
         BorderStyle     =   0  'None
         Caption         =   "����ѡ��"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   360
         TabIndex        =   7
         Top             =   4080
         Width           =   6135
         Begin VB.TextBox txtParaNum 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2160
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "ִ��������������Զ�ȡ�������Ĳ�������"
            Top             =   637
            Width           =   975
         End
         Begin VB.CheckBox chkParallel 
            Caption         =   "����ִ��"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "ִ��������������Զ�ȡ�������Ĳ�������"
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkOnln 
            Caption         =   "����ģʽ"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblParaNum 
            AutoSize        =   -1  'True
            Caption         =   "���ж�"
            Height          =   180
            Left            =   1560
            TabIndex        =   11
            Top             =   697
            Width           =   540
         End
      End
      Begin VB.Frame fraFKey 
         Enabled         =   0   'False
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   5520
         Width           =   6135
         Begin VB.OptionButton optDisable 
            Caption         =   "����Լ��"
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optDel 
            Caption         =   "ɾ��Լ��"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton optIdx 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtSql 
         Height          =   1065
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   6510
         Width           =   6615
      End
      Begin VB.Label lblSql 
         AutoSize        =   -1  'True
         Caption         =   "SQL����"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   6240
         Width           =   630
      End
      Begin VB.Label lblTip 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   7730
         Width           =   5565
      End
   End
   Begin VB.PictureBox pctLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkZlhis 
         Caption         =   "ֻ���ҵ���"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         ToolTipText     =   "�漰������ӱ�͸����Ϊҵ���"
         Top             =   7728
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtRange 
         Height          =   270
         Left            =   1080
         TabIndex        =   21
         Text            =   "100000"
         Top             =   7720
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   4320
         TabIndex        =   19
         Top             =   7680
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5295
         _cx             =   9340
         _cy             =   12726
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   75
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         SubtotalPosition=   0
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
      Begin VB.Label lblRange2 
         AutoSize        =   -1  'True
         Caption         =   "�еı�"
         Height          =   180
         Left            =   1800
         TabIndex        =   22
         Top             =   7770
         Width           =   540
      End
      Begin VB.Label lblRange1 
         AutoSize        =   -1  'True
         Caption         =   "ֻ������"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   7770
         Width           =   900
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         Caption         =   "���ȱʧ�������"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmIdxInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public Sub ShowMe()
    Me.Show
End Sub

    
Private Sub chkOnln_Click()
    Call GetSql
End Sub

Private Function Checkindex(ByVal strIndexName As String) As Boolean
'���ܣ����ָ���������Ƿ����
    Dim rstmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From DBA_Indexes Where Index_Name = [1] and Owner = [2]"
    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strIndexName, ".")(1), Split(strIndexName, ".")(0))
    
    Checkindex = rstmp.RecordCount > 0
    Exit Function
errH:
    Call ErrCenter(strSql)
End Function


Private Sub cmdExecute_Click()
    Dim intOldRow As Integer
    Dim strIdx As String, strFkey As String, strQuery As String
    
    On Error GoTo errH
    '������Ϣ
    If optIdx Then
        strQuery = "��ȷ��Ҫ����������" & optIdx.Tag & "����" & vbCrLf & vbCrLf & _
                    IIf(InStr(vsfGrid.TextMatrix(vsfGrid.Row, vsfGrid.ColIndex("����ֶ�")), ",") > 0, "���������������������������磺�ֶ�˳��ͬ���������ɾ��������Ӱ��ñ��д�����ܡ�" & vbCrLf, "") & _
                     IIf(chkOnln.Value = 0, "����û��ѡ�����ߴ����������ڼ䲻�ܶԸñ�����κ�д�����", "����ϴ��������ȽϺ�ʱ") & ",������ҵ������ڼ���С�"
    Else
        strQuery = "��ȷ��Ҫ" & IIf(optDel, "ɾ��", "����") & "Լ����" & optFKey.Tag & "����" & _
                        IIf(vsfGrid.Cell(flexcpText, vsfGrid.Row, vsfGrid.ColIndex("��������")) = "", "", vbCrLf & vbCrLf & "��Լ���м�����������ȷ�����Ķ���������˵����")
    End If
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then: Exit Sub
    
    gstrSQL = TrimEx(UCase(txtSql.Text))
    If Not (gstrSQL Like "CREATE INDEX*" Or gstrSQL Like "ALTER TABLE * CONSTRAINT*") Then
        MsgBox "ֻ����ִ���������������ɾ����ͣ�õ����"
        Exit Sub
    End If
    
    Call SetCmdEnable(False)
    intOldRow = vsfGrid.Row
    strIdx = optIdx.Tag
    strFkey = optFKey.Tag
    
    If gstrSQL Like "CREATE INDEX*" Then
        If Checkindex(strIdx) Then
            gcnOracle.Execute "Drop Index " & strIdx
        End If
    End If
    gcnOracle.Execute gstrSQL
    
    With vsfGrid
        .RemoveItem intOldRow
        If intOldRow >= .Rows - .FixedRows Then '��֤ѡ���е�λ�ò���
            .Select .Rows - .FixedRows, 1
        Else
            .Select intOldRow, 1
        End If
        .TopRow = .Row
        If .Row = intOldRow Then
            Call GetSql
        End If
    End With
    
    If optFKey.Value Then
        If optDel.Value Then
            lblTip.Caption = "ɾ����� " & strFkey & " �ɹ� ��"
        Else
            lblTip.Caption = "ͣ����� " & strFkey & " �ɹ� ��"
        End If
    Else
        lblTip.Caption = "�������� " & strIdx & " �ɹ� ��"
    End If
    
    Call SetCmdEnable(True)
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    lblTip.Caption = Err.Description

    Call SetCmdEnable(True)
End Sub

Private Sub cmdRefresh_Click()
    Dim intOldRow As Integer
    
    intOldRow = vsfGrid.Row
    Call SetCmdEnable(False)
    lblTip.Caption = "����ˢ�����ݣ����Եȡ�"
    lblTip.Refresh
    
    With vsfGrid
        .Rows = .FixedRows
        Call LoadGrid
        If intOldRow > .Rows - .FixedRows Then
            .Select .Rows - .FixedRows, 1
        Else
            .Select intOldRow, 1
        End If
        .TopRow = .Row
    End With
    Call SetCmdEnable(True)
    lblTip.Caption = ""
    
End Sub

Private Sub Form_load()
    Dim strCol As String
    
    '��ʼ����񣬼������ݣ����ÿؼ�������
    strCol = "������,1485,1;�ӱ�,1485,1;�������,2145,1;����ֶ�,1555,1;����,1705,1;��������,1300,1"
    Call InitTable(vsfGrid, strCol)
    With vsfGrid
        .Editable = flexEDNone
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .AutoSizeMode = flexAutoSizeColWidth
    End With
    
    Call LoadGrid
    
    With vsfGrid
        If .Rows > 2 Then
            vsfGrid.Select 1, 1
        End If
    End With
    
    optIdx.Value = True
    chkOnln.Value = 1
    optDel.Value = True

    txtParaNum.ToolTipText = "��ǰCPU��" & gintCpuCount & "�������鲢�ж� " & gintCpuAdvise & "����ж� " & gintCpuMax
    txtParaNum.Text = gintCpuAdvise
    
    lblRange1.Visible = Not gblnIsZlhis
    txtRange.Visible = Not gblnIsZlhis
    lblRange2.Visible = Not gblnIsZlhis
    chkZlhis.Visible = gblnIsZlhis
End Sub

Private Sub Form_Resize()

    '�������λ�ô�С
    pctLeft.Height = Me.ScaleHeight
    pctLeft.Width = Me.ScaleWidth - pctRight.Width - 25
    
End Sub

Private Sub optFKey_Click()
    Call SetEnableFra
    Call GetSql
End Sub

Private Sub optIdx_Click()
    '��������
    Call SetEnableFra
    Call GetSql
End Sub

Private Sub optDel_Click()
    Call GetSql
End Sub

Private Sub optDisable_Click()
    Call GetSql
End Sub

Private Sub chkParallel_Click()
    txtParaNum.Enabled = fraIdx.Enabled And chkParallel.Value
    Call GetSql
End Sub

Private Sub SetEnableFra()
'���ܣ��޸�ѡ��͸�ѡ��Ŀ�����
    fraIdx.Enabled = optIdx.Value And Not fraIdx.Enabled
    fraFKey.Enabled = optFKey.Value And Not fraFKey.Enabled
     
    chkOnln.Enabled = fraIdx.Enabled
    chkParallel.Enabled = fraIdx.Enabled
    lblParaNum.Enabled = fraIdx.Enabled
    txtParaNum.Enabled = fraIdx.Enabled And chkParallel.Value
    
    optDel.Enabled = fraFKey.Enabled
    optDisable.Enabled = fraFKey.Enabled
    
End Sub

Private Sub LoadGrid()
'���ܣ� ��ʼ����񣬼��ر�����ݡ�
    Dim rsData As ADODB.Recordset, i As Integer
    Dim strTblRange As String
    
    On Error GoTo errH
    
    If gblnIsZlhis Then
        '�Ƿ���Zltables���ű�
        If gblnHasZltables Then
            strTblRange = " c.Table_Name Not In (Select ���� From zlBaseCode) " & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And b.Table_Name In (Select ���� From zlTables Where ���� In ('B1','B2','B3','C1','C2','C3') )", "") & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And c.Table_Name In (Select ���� From Zltables Where ���� In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3'))", "") & vbNewLine
        Else
            strTblRange = " c.Table_Name Not In (Select ���� From zlBaseCode) " & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And b.Table_Name In (Select ���� From zlBakTables  " & IIf(gblnHasBigtables, "Union All Select ���� From Zlbigtables )", ")"), "") & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And c.Table_Name In (Select ���� From zlBakTables  " & IIf(gblnHasBigtables, "Union All Select ���� From Zlbigtables )", ")"), "") & vbNewLine
        End If
    Else
        strTblRange = "b.Table_Name in (Select Table_Name From Dba_Tables Where Num_Rows > " & Val(txtRange.Text) & ")" & vbNewLine
    End If
    
    '��ѯ������������ӱ�ȱʧ����������
    'Child_Table-�ӱ�   Foreign_Key-���  Columns-�ӱ�������  Main_Table-���� Delete_Rule-���ɾ������
    gstrSQL = "Select Main_Table, Child_Table, Foreign_Key, Columns, Delete_Rule, Owner" & vbNewLine & _
                        "From (Select c.Table_Name As Main_Table, b.Table_Name As Child_Table, b.Constraint_Name As Foreign_Key," & vbNewLine & _
                        "              f_List2str(Cast(Collect(a.Column_Name Order By Position) As t_Strlist), ',', 1) As Columns, b.Delete_Rule," & vbNewLine & _
                        "              b.Owner" & vbNewLine & _
                        "       From Dba_Cons_Columns A, Dba_Constraints B, Dba_Constraints C" & vbNewLine & _
                        "       Where a.Constraint_Name = b.Constraint_Name And b.Status = 'ENABLED' And b.Constraint_Type = 'R' And" & vbNewLine & _
                        "             b.r_Constraint_Name <> '���ű�_PK' And b.r_Constraint_Name = c.Constraint_Name And b.r_owner=c.owner  And A.Owner = B.Owner And" & vbNewLine & _
                        strTblRange & _
                        "       Group By c.Table_Name, b.Table_Name, b.Delete_Rule, b.Constraint_Name, b.r_Constraint_Name, b.Owner) A " & vbNewLine & _
                        "Where Not Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From (Select Table_Name, Index_Name, Index_Owner," & vbNewLine & _
                        "                     f_List2str(Cast(Collect(Column_Name Order By Column_Position) As t_Strlist), ',', 1) As Columns" & vbNewLine & _
                        "              From Dba_Ind_Columns" & vbNewLine & _
                        "              Group By Table_Name, Index_Name, Index_Owner) B" & vbNewLine & _
                        "       Where a.Owner = b.Index_Owner And b.Table_Name = a.Child_Table And Instr(b.Columns, a.Columns) =1)" & vbNewLine
                        
    Set rsData = OpenSQLRecord(gstrSQL, Me.Caption)
    If rsData.RecordCount = 0 Then
        Call ClearVsf(vsfGrid, "��ǰ����û������ӱ�ȱʧ������")
        Exit Sub
    End If
    
    With vsfGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsData.RecordCount + .FixedRows
        
        i = 1
        While Not rsData.EOF
            .TextMatrix(i, .ColIndex("������")) = rsData!Owner
            .TextMatrix(i, .ColIndex("�ӱ�")) = rsData!Child_Table
            .TextMatrix(i, .ColIndex("�������")) = rsData!Foreign_Key
            .TextMatrix(i, .ColIndex("����ֶ�")) = rsData!Columns
            .TextMatrix(i, .ColIndex("����")) = rsData!Main_Table
            
            If rsData!Delete_Rule = "CASCADE" Then
                .TextMatrix(i, .ColIndex("��������")) = "����"
            ElseIf rsData!Delete_Rule = "SET NULL" Then
                .TextMatrix(i, .ColIndex("��������")) = "���"
            Else
                .TextMatrix(i, .ColIndex("��������")) = ""
            End If
            
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = BackAlterNate_��ɫ
            Else
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = Back_��ɫ
            End If
            
            i = i + 1
            rsData.MoveNext
        Wend
        .AutoSize .ColIndex("�ӱ�"), .ColIndex("��������"), False
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    Call SetCmdEnable(True)
End Sub

Private Sub pctLeft_Resize()
    On Error Resume Next
    vsfGrid.Width = pctLeft.Width - vsfGrid.Left
    vsfGrid.Height = pctLeft.Height - cmdRefresh.Height - lblPrompt.Height - 280
    
    cmdRefresh.Top = vsfGrid.Top + vsfGrid.Height + 60
    cmdRefresh.Left = vsfGrid.Left + vsfGrid.Width - cmdRefresh.Width
    
    lblRange1.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - lblRange1.Height / 2
    lblRange2.Top = lblRange1.Top
    txtRange.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - txtRange.Height / 2
    
    lblRange1.Left = vsfGrid.Left
    txtRange.Left = lblRange1.Left + lblRange1.Width + 45
    lblRange2.Left = txtRange.Left + txtRange.Width + 45
    
    chkZlhis.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - chkZlhis.Height / 2 + 15
    chkZlhis.Left = cmdRefresh.Left - chkZlhis.Width
End Sub

Private Sub pctRight_Resize()
    
    On Error Resume Next
    txtSql.Height = vsfGrid.Top + vsfGrid.Height - txtSql.Top
    cmdExecute.Top = txtSql.Height + txtSql.Top + 60
    
    lblTip.Top = cmdExecute.Top + cmdRefresh.Height - lblTip.Height
    lblTip.Width = pctRight.Width - (cmdExecute.Left + cmdExecute.Width + 60)
End Sub

Private Sub txtParaNum_Change()
    Call GetSql
End Sub

Private Sub txtParaNum_KeyPress(KeyAscii As Integer)
    Call OnlyIntCK(KeyAscii)
    If Val(txtParaNum.Text & Chr(KeyAscii)) > gintCpuCount And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRange_KeyPress(KeyAscii As Integer)
    Call OnlyIntCK(KeyAscii)
End Sub

Private Sub txtSql_KeyPress(KeyAscii As Integer)
    Call OnlyStrChnCK(KeyAscii, " ", "_", Chr(3), Chr(22))
End Sub

Private Function GetTbSpace(ByVal strTbName As String, ByVal strOwner As String) As String
'���ܣ����ݱ�����ȡ�������ڵı�ռ�
'������ strTbName - ����
    Dim rstmp As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "select a.table_name , a.tablespace_name ,count (b.index_name) index_Nums  ,b.tablespace_name Index_tbs" & vbNewLine & _
                    "from dba_tables  a , dba_indexes b" & vbNewLine & _
                    "where a.table_name = b.table_name(+)  and a.table_name = [1] and a.owner = [2]" & vbNewLine & _
                    "and a.temporary = 'N' and a.partitioned = 'NO'   and a.owner = b.owner" & vbNewLine & _
                    "group by  a.table_name,  a.tablespace_name ,b.tablespace_name" & vbNewLine & _
                    "order by index_Nums desc"

    Set rstmp = OpenSQLRecord(gstrSQL, Me.Caption, strTbName, strOwner)
    
    If rstmp.RecordCount = 0 Then Exit Function
    GetTbSpace = IIf(rstmp!index_Nums = 0, rstmp!Tablespace_Name, rstmp!Index_tbs)
    
    Exit Function
errH:
    Call ErrCenter(gstrSQL)

End Function

Private Sub GetSql()
'���� �����ݽ�����ѡ���޸�SQLָ��
'�ؼ�ѡ����ڶ�ӦTAG�У�ʵ���������
'create index  [����]  on [����]([��]) tablespace [��ռ�]  nologging [���ж�] [����];
'alter table [����] [����] constraint [���]
    Dim strTbName As String, strCols As String, srtFkey As String
    Dim strTbSpace As String, strOwner As String
    
    With vsfGrid
      
        If .Rows < 2 Or .Row = 0 Then txtSql.Text = "": Exit Sub '��ֹ����Խ��
        If .TextMatrix(1, 0) = "��ǰ����û������ӱ�ȱʧ������" Then
            txtSql.Text = ""
            cmdExecute.Enabled = False
            Exit Sub
        Else
            cmdExecute.Enabled = True
        End If
        strTbName = .TextMatrix(.Row, .ColIndex("�ӱ�"))
        strCols = .TextMatrix(.Row, .ColIndex("����ֶ�"))
        srtFkey = .TextMatrix(.Row, .ColIndex("�������"))
        strOwner = .TextMatrix(.Row, .ColIndex("������"))
        strTbSpace = GetTbSpace(strTbName, strOwner)
    End With
    
    If optFKey Then  'ɾ�����
        optFKey.Tag = srtFkey
        optDel.Tag = IIf(optDel.Value, " Drop", " Disable")
        txtSql.Text = "Alter Table " & strOwner & "." & strTbName & optDel.Tag & " Constraint " & srtFkey
    Else    '��������
        If InStr(1, strCols, ",") > 0 Then
            optIdx.Tag = strOwner & "." & strTbName & "_IX_" & Mid(strCols, 1, InStr(1, strCols, ",") - 1)
        Else
            optIdx.Tag = strOwner & "." & strTbName & "_IX_" & strCols
        End If
        chkParallel.Tag = IIf(chkParallel.Value = 1, " Parallel " & txtParaNum.Text & " ", " ")
        chkOnln.Tag = IIf(chkOnln.Value = 1, "online ", " ")
        txtSql.Text = "Create Index " & optIdx.Tag & " On " & strOwner & "." & strTbName & "(" & strCols & ") Tablespace " & strTbSpace & " nologging" & _
                             chkParallel.Tag & chkOnln.Tag
    End If
    txtSql.Tag = txtSql.Text
        
End Sub


Private Sub vsfGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow = 0 Then: Exit Sub   '��ֹˢ��ʱ��
    If txtSql.Tag <> txtSql.Text Then
        If MsgBox("��ǰSQLָ���Ѿ��������ģ��л��������޸Ľ���ʧ��" & vbCrLf & vbCrLf & "�Ƿ��л���", _
            vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub vsfGrid_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call GetSql
End Sub

Private Sub SetCmdEnable(ByVal blnEnable As Boolean)
'���ܣ� ���ð�ť�����Ժ͹����ʽ
    cmdRefresh.Enabled = blnEnable
    cmdExecute.Enabled = cmdRefresh.Enabled
    If blnEnable Then
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbArrowHourglass
    End If
End Sub

