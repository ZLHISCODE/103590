VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmNoneedExecBloodSelector 
   Caption         =   "����ִ��ѪҺѡ��"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5850
   Icon            =   "frmNoneedExecBloodSelector.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5850
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdEdit 
      Caption         =   "��д(&E)"
      Height          =   855
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   4800
      TabIndex        =   5
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Default         =   -1  'True
      Height          =   330
      Left            =   3720
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtReson 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3360
      Width           =   4860
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5610
      _cx             =   9895
      _cy             =   4154
      Appearance      =   2
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
      BackColorSel    =   16444122
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
      Rows            =   9
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmNoneedExecBloodSelector.frx":6852
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ͳһ��д�ѹ�ѡѪҺ�ĸ���ԭ��"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3135
      Width           =   2700
   End
   Begin VB.Label lbl 
      Caption         =   "��ע���������������Ѫ��Ӧ,�ɽ���ҽ����δִ����ɵ�ѪҺ����Ϊ����ִ��״̬"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3690
   End
End
Attribute VB_Name = "frmNoneedExecBloodSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnShow As Boolean
Private mblnFinish As Boolean
Private mclsVsf As clsVsf
Private mlngҽ��ID As Long
Private mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal frmParent As Object, ByVal lngҽ��ID As Long, Optional blnFinish As Boolean) As Boolean
    
    mblnOK = False
    mblnFinish = False

    mlngҽ��ID = lngҽ��ID
    mblnShow = False
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    blnFinish = mblnFinish
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Dim i As Integer
    
    With vsExec
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked Then
                .TextMatrix(i, .ColIndex("ִ��ժҪ")) = Trim(txtReson.Text)
            End If
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim mrsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnUnload As Boolean
    
    On Error GoTo ErrHand
    '��ʼ�����
    Call InitTable
    
    '����ѪҺִ�����
    Call LoadExecVsf
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTable()
'����ʼ��
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsExec, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTBoolean, "", "ѡ��", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "ID", , True, False, False, True)  '�շ�ID
        Call .AppendColumn("״̬", 800, flexAlignLeftCenter, flexDTString, , "ѪҺ״̬") '����ִ��״̬
        Call .AppendColumn("Ѫ�����", 1200, flexAlignLeftCenter, flexDTString, , "Ѫ�����", , False, False)
        Call .AppendColumn("ѪҺ����", 1300, flexAlignLeftCenter, flexDTString, , "ѪҺ����", , False, False)
        Call .AppendColumn("����ԭ��", 1300, flexAlignLeftCenter, flexDTString, , "ִ��ժҪ", , False, False)
        .AppendRows = False
    End With
        vsExec.ExplorerBar = flexExNone
End Sub

Private Sub LoadExecVsf()
    '���ܣ�����ѪҺִ�����
    Dim i As Integer, intRow As Integer
    Dim arrName, arrKey, arrColWidth
    Dim strSQL As String
    

    On Error GoTo ErrHand
    strSQL = "SELECT a.Id, decode(h.ִ��״̬,4,1,0) ѡ��, a.Ѫ�����, h.ִ��״̬, h.����״̬, h.ִ��ժҪ," & vbNewLine & _
                "       Decode(Nvl(h.ִ��״̬, 0), 0, '�ѽ���', 4, '����ִ��') ѪҺ״̬, e.���� AS ѪҺ����" & vbNewLine & _
                "FROM �շ���ĿĿ¼ e, ѪҺƷ�� k, ѪҺ��� l, ѪҺ�շ���¼ a, ѪҺ���ͼ�¼ h, ѪҺ��Ѫ��¼ b" & vbNewLine & _
                "WHERE e.Id = a.ѪҺid AND k.Ʒ��id = l.Ʒ��id AND l.���id = a.ѪҺid AND a.Id = h.�շ�id AND h.�䷢id = b.Id AND h.����״̬ = 1 AND" & vbNewLine & _
                "      h.ִ��״̬ IN (0, 4) AND Nvl(a.��Ѫ״̬, 0) = 2 AND b.����id = [1]" & vbNewLine & _
                "ORDER BY h.ִ�з���, a.��Ѫ����, a.���"
    Set mrsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������ִ�е�ѪҺ", mlngҽ��ID)
    If mrsTmp.RecordCount = 0 Then
        Call MsgBox("���ѽ��ջ�����ִ�е�ѪҺ��", vbInformation, "�������")
        Unload Me
        Exit Sub
    End If
    Call mclsVsf.LoadGrid(mrsTmp, "", True)
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, strSQL As String
    Dim arrSQL, i As Integer
    Dim intStatus As Integer
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With vsExec
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 And .TextMatrix(i, .ColIndex("ִ��ժҪ")) = "" Then
            MsgBox "��" & i & "��ѪҺδ��д����ԭ������д���ٽ��б��棡", vbInformation, "�������"
            Exit Function
        End If
    Next
    arrSQL = Array()
        For i = 1 To .Rows - 1
            mrsTmp.Filter = "ID = " & .TextMatrix(i, .ColIndex("ID"))
            intStatus = mrsTmp!ִ��״̬
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = vbChecked And .TextMatrix(i, .ColIndex("ѪҺ״̬")) <> "����ִ��" Then
                intStatus = 4
            ElseIf .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 2 And .TextMatrix(i, .ColIndex("ѪҺ״̬")) = "����ִ��" Then
                strSQL = _
                    " SELECT 1 FROM ѪҺִ�м�¼ WHERE �շ�id = [1] AND ROWNUM < 2"
                Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "ѪҺִ�м�¼��ȡ", .TextMatrix(i, .ColIndex("ID")))
                If rsData.RecordCount > 0 Then
                    intStatus = 1
                Else
                    intStatus = 0
                End If
            End If
            If mrsTmp.RecordCount <> 0 Then
                If intStatus <> mrsTmp!ִ��״̬ Or .TextMatrix(i, .ColIndex("ִ��ժҪ")) <> mrsTmp!ִ��ժҪ & "" Then
                    strSQL = "Zl_ѪҺ���ͼ�¼_����ִ��(" & .TextMatrix(i, .ColIndex("ID")) & "," & intStatus & "," & "'" & _
                                .TextMatrix(i, .ColIndex("ִ��ժҪ")) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                End If
            End If
            mrsTmp.Filter = ""
        Next
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            If CStr(arrSQL(i)) <> "" Then
                Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            End If
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End With
    SaveData = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txtReson_KeyPress(KeyAscii As Integer)
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsExec_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsExec.ColIndex("ѡ��") And Col <> vsExec.ColIndex("ִ��ժҪ") Then Cancel = True
    If (Col = vsExec.ColIndex("ִ��ժҪ") And vsExec.Cell(flexcpChecked, Row, vsExec.ColIndex("ѡ��")) = 2) Then Cancel = True
End Sub

Private Sub vsExec_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsExec_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsExec.ColIndex("ѡ��") And vsExec.Cell(flexcpChecked, Row, vsExec.ColIndex("ѡ��")) = vbChecked Then
        vsExec.TextMatrix(Row, vsExec.ColIndex("ִ��ժҪ")) = ""
    End If
End Sub
