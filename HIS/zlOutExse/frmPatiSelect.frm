VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPatiSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ѡ��"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7875
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   3
      Top             =   4395
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4875
      TabIndex        =   2
      Top             =   4395
      Width           =   1150
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   4245
      Left            =   2670
      TabIndex        =   1
      Top             =   15
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12640511
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDept 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12640511
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlngPatient As Long
Private mrsPati As New ADODB.Recordset

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mshPati.Rows > 1 And mshPati.TextMatrix(1, 0) <> "" Then
        mlngPatient = Val(mshPati.TextMatrix(mshPati.Row, 0))
        Unload Me
    End If
End Sub

Private Sub mshDept_EnterCell()
    Dim i As Long, j As Long, strSQL As String
    
    Me.Refresh
    mshPati.Clear
    If mshDept.RowData(mshDept.Row) = 0 Then Exit Sub
    
    On Error GoTo errHandle

    '��ǰ��Ժ����:Ŀǰ�������������۲���,����סԺ���۲���
    strSQL = _
        " Select A.����ID,A.סԺ��,A.����,A.��ǰ���� as ��λ,A.�Ա�,B.�ѱ�,'��' as ��Ժ,B.����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID" & _
        " And A.��ҳID=B.��ҳID And Nvl(B.��ҳID,0)<>0" & _
        " And A.��Ժ=1 And Nvl(B.��������,0)<>1" & _
        " And B.��Ժ����ID+0=[1]" & _
        " Order by A.סԺ�� Desc"
    Screen.MousePointer = 11
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mshDept.RowData(mshDept.Row)))
    
    With mshPati
        .Redraw = False
        If Not mrsPati.EOF Then
            Set .Recordset = mrsPati
            .ColWidth(0) = 800
            .ColWidth(1) = 800
            .ColWidth(2) = 850
            .ColWidth(3) = 600
            .ColWidth(4) = 500
            .ColWidth(5) = 800
            .ColWidth(6) = 500
            .ColWidth(7) = 0
            .ColAlignment(4) = 4
            .ColAlignment(5) = 1
            .ColAlignment(6) = 4
            
            '����ҽ�����˵���ɫ
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 7)) <> 0 Then
                    .Row = i
                    For j = 0 To .COLS - 1
                        .Col = j
                        .CellForeColor = vbRed
                    Next
                End If
            Next
            .Row = 1: .Col = 0: .ColSel = .COLS - 1
            Call mshPati_EnterCell
        Else
            .Clear
            .ClearStructure
            .Rows = 2
            .COLS = 2
        End If
        .Redraw = True
    End With
    
    For i = 0 To mshPati.COLS - 1
        mshPati.ColAlignmentFixed(i) = 4
    Next
    mshPati.RowHeight(0) = 320
    mshPati.Row = 1: mshPati.TopRow = 1
    mshPati.Col = 0: mshPati.ColSel = mshPati.COLS - 1
    
    Screen.MousePointer = 0
    
    If Not mrsPati.EOF Then
        If Visible Then mshPati.SetFocus
    Else
        If Visible Then mshDept.SetFocus
    End If
        

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mshDept_GotFocus()
    mshDept.BackColorSel = &HC0E0FF
    mshPati.BackColorSel = &HC0C0C0
End Sub

Private Sub mshDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mshPati_KeyDown(KeyCode, Shift)
End Sub

Private Sub mshPati_DblClick()
    cmdOK_Click
End Sub

Private Sub mshPati_EnterCell()
    If mshPati.CellForeColor = vbRed Then
        mshPati.ForeColorSel = vbRed
    Else
        mshPati.ForeColorSel = mshDept.ForeColorSel
    End If
End Sub

Private Sub mshPati_GotFocus()
    mshDept.BackColorSel = &HC0C0C0
    mshPati.BackColorSel = &HC0E0FF
End Sub

Private Sub mshPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    mshPati.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
            
    On Error GoTo errHandle
    mlngPatient = 0
    
    mshDept.TextMatrix(0, 0) = "����"
    mshDept.TextMatrix(0, 1) = "����"
    mshDept.Rows = 2: mshDept.COLS = 2
    mshDept.ColAlignmentFixed(0) = 4
    mshDept.ColAlignmentFixed(1) = 4
    mshDept.ColAlignment(0) = 1
    mshDept.ColAlignment(1) = 1
    mshDept.ColWidth(0) = 830
    mshDept.ColWidth(1) = 1500
    mshDept.Row = 1
        
    'ȡ�в��˵�סԺ����:Ŀǰ�������������۲���,����סԺ���۲���
    strSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
            "From ���ű� a, ��������˵�� b" & vbNewLine & _
            "Where a.Id = b.����id And b.�������� = '�ٴ�' And b.������� In (1, 2, 3)" & vbNewLine & _
            " And Exists(Select 'X' From ��λ״����¼ x Where x.����id Is Not Null Group By x.����id Having x.����id = a.Id)" & vbNewLine & _
            " And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "Order By a.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    With rsTmp
        If Not .EOF Then
            mshDept.Rows = rsTmp.RecordCount + 1
            For i = 1 To .RecordCount
                mshDept.TextMatrix(i, 0) = !����
                mshDept.TextMatrix(i, 1) = !����
                mshDept.RowData(i) = !ID
                If UserInfo.����ID = !ID Then mshDept.Row = i 'ֱ����������
                .MoveNext
            Next
        End If
    End With
    
    mshDept.Col = 0: mshDept.ColSel = mshDept.COLS - 1
    mshDept.TopRow = mshDept.Row
    Call mshDept_EnterCell
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        If mshDept.RowData(mshDept.Row) = 0 Then Exit Sub
        If KeyCode = vbKeyLeft Then
            If mshDept.Row - 1 >= 1 Then mshDept.Row = mshDept.Row - 1
        ElseIf KeyCode = vbKeyRight Then
            If mshDept.Row + 1 <= mshDept.Rows - 1 Then
                mshDept.Row = mshDept.Row + 1
            End If
        End If
        mshDept.Col = 0: mshDept.ColSel = mshDept.COLS - 1
        If mshDept.CellTop + mshDept.CellHeight > mshDept.Height - 300 Then mshDept.TopRow = mshDept.TopRow + 1
        Call mshDept_EnterCell
        mshPati.Col = 0: mshPati.ColSel = mshPati.COLS - 1
    End If
End Sub
