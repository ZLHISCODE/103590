VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAdviceFreqTime 
   Caption         =   "ִ��ʱ�䷽��"
   ClientHeight    =   4875
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7740
   Icon            =   "frmAdviceFreqTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7740
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   6480
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "����(&I)"
      Height          =   350
      Left            =   6495
      TabIndex        =   3
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6495
      TabIndex        =   2
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6495
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdVaild 
      Caption         =   "�Ϸ���(&V)"
      Height          =   350
      Left            =   6495
      TabIndex        =   4
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6495
      TabIndex        =   5
      Top             =   4350
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit mshTime 
      Height          =   3990
      Left            =   105
      TabIndex        =   0
      Top             =   765
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7038
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdviceFreqTime.frx":030A
      Height          =   560
      Left            =   780
      TabIndex        =   6
      Top             =   90
      Width           =   5670
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   165
      Picture         =   "frmAdviceFreqTime.frx":03A3
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAdviceFreqTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrCode As String
Private mintƵ�ʴ��� As Integer
Private mintƵ�ʼ�� As Integer
Private mstr�����λ As String
Private mint���÷�Χ As Integer '1-��ҽ,2-��ҽ
Private mblnChange As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Long
    
    If mshTime.Row >= mshTime.MsfObj.FixedRows Then
        If mshTime.Rows = mshTime.MsfObj.FixedRows + 1 Then
            For i = 1 To mshTime.Cols - 1
                mshTime.TextMatrix(mshTime.MsfObj.FixedRows, i) = ""
            Next
        Else
            Call mshTime.MsfObj.RemoveItem(mshTime.Row)
        End If
    End If
    Call AdjustOrder
    mshTime.SetFocus
End Sub

Private Sub cmdInsert_Click()
    mshTime.MsfObj.AddItem "", mshTime.Row
    mshTime.Col = 1
    Call AdjustOrder
    mshTime.SetFocus
End Sub

Private Function CheckValid(lngRow As Long, lngCol As Long, strErr As String, Optional arrSql As Variant) As Boolean
    Dim strTime As String, i As Long, j As Long
    
    '�ȼ��Ϸ���
    arrSql = Array()
    For i = mshTime.MsfObj.FixedRows To mshTime.Rows - 1
        If mshTime.TextMatrix(i, 1) <> "" Then
            For j = 2 To mshTime.Cols - 1
                If mshTime.TextMatrix(i, j) = "" Then
                    strErr = "���Ϊ " & mshTime.TextMatrix(i, 0) & " ��ʱ�䷽������������ݲ�������"
                    lngRow = i: lngCol = j: Exit For
                End If
            Next
            If j <= mshTime.Cols - 1 Then Exit For
            
            strTime = ""
            If mstr�����λ = "��" Or mstr�����λ = "��" And mintƵ�ʼ�� > 1 Then
                For j = 2 To mshTime.Cols - 1 Step 2
                    strTime = strTime & "-" & mshTime.TextMatrix(i, j) & "/" & mshTime.TextMatrix(i, j + 1)
                Next
            ElseIf mstr�����λ = "Сʱ" Or mstr�����λ = "��" And mintƵ�ʼ�� = 1 Then
                For j = 2 To mshTime.Cols - 1
                    strTime = strTime & "-" & mshTime.TextMatrix(i, j)
                Next
            End If
            strTime = Mid(strTime, 2)
            
            If Not ExeTimeValid(strTime, lngCol, strErr) Then
                If lngCol = 0 Then
                    lngCol = 2
                    strErr = "���Ϊ " & mshTime.TextMatrix(i, 0) & " ��ʱ�䷽����" & strErr
                Else
                    strErr = "���Ϊ " & mshTime.TextMatrix(i, 0) & " ��ʱ�䷽����" & mshTime.TextMatrix(0, lngCol) & strErr
                End If
                lngRow = i: Exit For
            End If
            
            If zlCommFun.ActualLen(strTime) > 50 Then
                strErr = "ʱ�䷽������̫���������������Ƶ�ʴ���̫�����£����ȶԸ�Ƶ����Ŀ���ʵ�������"
                lngRow = i: Exit For
            End If
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "ZL_����Ƶ��ʱ��_Insert('" & mstrCode & "'," & mshTime.TextMatrix(i, 0) & "," & _
                "'" & strTime & "'," & IIf(mshTime.RowData(i) = 0, "NULL", mshTime.RowData(i)) & ")"
        Else
            If i <> mshTime.Rows - 1 Then
                strErr = "���Ϊ " & mshTime.TextMatrix(i, 0) & " ��ʱ�䷽������������ݲ�������"
                lngRow = i: lngCol = 1: Exit For
            Else
                For j = 2 To mshTime.Cols - 1
                    If mshTime.TextMatrix(i, j) <> "" Then
                        strErr = "���Ϊ " & mshTime.TextMatrix(i, 0) & " ��ʱ�䷽������������ݲ�������"
                        lngRow = i: lngCol = 1: Exit For
                    End If
                Next
                If j <= mshTime.Cols - 1 Then Exit For
            End If
        End If
    Next
    CheckValid = Not (i <= mshTime.Rows - 1)
End Function

Private Sub cmdOK_Click()
    Dim arrSql As Variant, strErr As String
    Dim lngRow As Long, lngCol As Long, i As Long
    
    If Not CheckValid(lngRow, lngCol, strErr, arrSql) Then
        mshTime.Row = lngRow: mshTime.Col = lngCol
        Call mshTime_EnterCell(lngRow, lngCol)
        If lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2 < mshTime.MsfObj.FixedRows Then
            mshTime.MsfObj.TopRow = mshTime.MsfObj.FixedRows
        Else
            mshTime.MsfObj.TopRow = lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2
        End If
        If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName
        mshTime.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure("ZL_����Ƶ��ʱ��_Delete('" & mstrCode & "')", Me.Caption)
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    mblnChange = False
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdVaild_Click()
    Dim strErr As String, lngRow As Long, lngCol As Long
    
    If Not CheckValid(lngRow, lngCol, strErr) Then
        mshTime.Row = lngRow: mshTime.Col = lngCol
        Call mshTime_EnterCell(lngRow, lngCol)
        If lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2 < mshTime.MsfObj.FixedRows Then
            mshTime.MsfObj.TopRow = mshTime.MsfObj.FixedRows
        Else
            mshTime.MsfObj.TopRow = lngRow - mshTime.Height \ mshTime.RowHeight(0) \ 2
        End If
        If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName
        mshTime.SetFocus
        Exit Sub
    Else
        MsgBox "��������������ȷ��", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If mshTime.TxtVisible Then
            mshTime.Text = "": mshTime.TxtVisible = False: mshTime.SetFocus
        Else
            Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    RestoreWinState Me, App.ProductName
    gblnOK = False
    mblnChange = False
        
    'Ƶ����Ŀ��Ϣ
    strSql = "Select * From ����Ƶ����Ŀ Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrCode)
    
    mintƵ�ʴ��� = Nvl(rsTmp!Ƶ�ʴ���, 0)
    mintƵ�ʼ�� = Nvl(rsTmp!Ƶ�ʼ��, 0)
    mstr�����λ = Nvl(rsTmp!�����λ)
    mint���÷�Χ = IIf(IsNull(rsTmp!���÷�Χ), 1, rsTmp!���÷�Χ)
    
    lblCaption.Caption = Replace(lblCaption.Caption, "XXXXX", rsTmp!����)
    lblCaption.Caption = Replace(lblCaption.Caption, "YYYYY", IIf(mint���÷�Χ = 1, "��ҩ;��", "��ҩ�÷�"))
                    
    '��ʾ����ʽ������
    Call ShowTimeScheme(mstrCode)
    
    '�б༭����
    mshTime.ColData(0) = 5
    mshTime.ColData(1) = 1
    For i = 2 To mshTime.Cols - 1
        mshTime.ColData(i) = 4
    Next
    mshTime.LocateCol = 1
    mshTime.PrimaryCol = 1

    mshTime.Col = 1
    mshTime.Row = mshTime.MsfObj.FixedRows
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With mshTime
        .Width = Me.ScaleWidth - .Left - cmdOK.Width - 350
        .Height = Me.ScaleHeight - .Top - 60
    End With
    
    cmdOK.Left = mshTime.Left + mshTime.Width + 200
    cmdCancel.Left = cmdOK.Left
    cmdVaild.Left = cmdOK.Left
    cmdInsert.Left = cmdOK.Left
    cmdHelp.Left = cmdOK.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("�����޸���������ݣ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    SaveWinState Me, App.ProductName
    mstrCode = ""
End Sub

Private Function ShowTimeScheme(ByVal str���� As String) As Boolean
'���ܣ����ݵ�ǰƵ����Ŀ��ʾ����ʱ�䷽����
'������str����=Ƶ����Ŀ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim arrTime As Variant
    
    On Error GoTo errH
    
    With mshTime.MsfObj
        .Clear
        .ClearStructure
        .FixedCols = 0: .FixedRows = 0
        .Rows = 0: .Cols = 0
        
        'Ƶ��ʱ�䷽��
        strSql = _
            "Select A.�������,A.ʱ�䷽��,A.��ҩ;��ID,B.����,B.����" & _
            " From ����Ƶ��ʱ�� A,������ĿĿ¼ B" & _
            " Where A.��ҩ;��ID=B.ID(+) And A.ִ��Ƶ��=[1]" & _
            " Order by �������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����)
        
        'ʱ�䷽����ͷ
        If mstr�����λ = "��" Or mstr�����λ = "��" And mintƵ�ʼ�� > 1 Then
            .Cols = 2 + mintƵ�ʴ��� * 2
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 2
            .FixedRows = 2
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "���": .TextMatrix(1, 0) = .TextMatrix(0, 0)
            .TextMatrix(0, 1) = IIf(mint���÷�Χ = 1, "��ҩ;��", "��ҩ�÷�"): .TextMatrix(1, 1) = .TextMatrix(0, 1)
            For i = 2 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "��" & ((i - 2) \ 2) + 1 & "��"
                .TextMatrix(0, i + 1) = .TextMatrix(0, i)
                If mstr�����λ = "��" Then
                    .TextMatrix(1, i) = "����"
                    .TextMatrix(1, i + 1) = "ʱ��"
                    .ColWidth(i) = 450
                    .ColWidth(i + 1) = 1000
                Else
                    .TextMatrix(1, i) = "��"
                    .TextMatrix(1, i + 1) = "ʱ��"
                    .ColWidth(i) = 300
                    .ColWidth(i + 1) = 1000
                End If
                .ColAlignment(i) = 4
                .ColAlignment(i + 1) = 1
            Next
        ElseIf mstr�����λ = "Сʱ" Or mstr�����λ = "��" And mintƵ�ʼ�� = 1 Then
            .Cols = 2 + mintƵ�ʴ���
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 1
            .FixedRows = 1
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = IIf(mint���÷�Χ = 1, "��ҩ;��", "��ҩ�÷�")
            For i = 2 To .Cols - 1
                .TextMatrix(0, i) = "��" & i - 1 & "��"
                .ColWidth(i) = 1000
                .ColAlignment(i) = 1
            Next
        End If
        .ColWidth(0) = 450
        .ColWidth(1) = 1800
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        For i = 0 To .Cols - 1
            .ColAlignmentFixed(i) = 4
        Next
        .MergeCells = flexMergeRestrictAll
        .MergeCol(0) = True: .MergeCol(1) = True
        .MergeRow(0) = True: .MergeRow(1) = True
        
        'ʱ������
        Call AdjustOrder
        For i = 1 To rsTmp.RecordCount
            .RowData(i + .FixedRows - 1) = IIf(IsNull(rsTmp!��ҩ;��ID), 0, rsTmp!��ҩ;��ID)
            .TextMatrix(i + .FixedRows - 1, 0) = rsTmp!�������
            .TextMatrix(i + .FixedRows - 1, 1) = IIf(IsNull(rsTmp!����), "<��ȷ��>", rsTmp!���� & "-" & rsTmp!����)
            
            arrTime = Split(rsTmp!ʱ�䷽��, "-")
            If mstr�����λ = "��" Or mstr�����λ = "��" And mintƵ�ʼ�� > 1 Then
                For j = 0 To mintƵ�ʴ��� - 1
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 2) = Split(arrTime(j), "/")(0)
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 3) = Split(arrTime(j), "/")(1)
                Next
            ElseIf mstr�����λ = "Сʱ" Or mstr�����λ = "��" And mintƵ�ʼ�� = 1 Then
                For j = 0 To mintƵ�ʴ��� - 1
                    .TextMatrix(i + .FixedRows - 1, j + 2) = arrTime(j)
                Next
            End If
            rsTmp.MoveNext
        Next
    End With
    
    ShowTimeScheme = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshTime_AfterAddRow(Row As Long)
    Call AdjustOrder(Row)
End Sub

Private Sub mshTime_AfterDeleteRow()
    Call AdjustOrder(mshTime.Row)
    mblnChange = True
End Sub

Private Sub AdjustOrder(Optional ByVal lngRow As Long)
    Dim i As Long
    
    If lngRow = 0 Then lngRow = mshTime.MsfObj.FixedRows
    
    For i = lngRow To mshTime.Rows - 1
        mshTime.TextMatrix(i, 0) = i - mshTime.MsfObj.FixedRows + 1
    Next
End Sub

Public Function ExeTimeValid(ByVal strTime As String, lngCol As Long, strErr As String) As Boolean
'���ܣ����ָ����ִ��ʱ���Ƿ�Ϸ�
'���أ�lngCol=����������
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then Exit Function
    
    If mstr�����λ = "��" Then
        '1/8:00-3/15:00-5/9:00��1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then
            strErr = "�����˷Ƿ��ַ���": Exit Function
        End If
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> mintƵ�ʴ��� Then
            strErr = "�����˷Ƿ��ַ���": Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then
                strErr = "�����˷Ƿ��ַ���"
                lngCol = i * 2 + 2: Exit Function
            End If
            
            '���ڲ���
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then
                strErr = "�����ڲ������벻��ȷ��"
                lngCol = i * 2 + 2: Exit Function
            End If
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then
                strErr = "�����ڱ����� 1-7 ֮�䡣"
                lngCol = i * 2 + 2: Exit Function
            End If
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then
                    strErr = "������������С����һ�ε���������"
                    lngCol = i * 2 + 2: Exit Function
                End If
            End If
            
            '����ʱ�䲿��
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then
                strErr = "�����˶��ʱ��ָ���"":""��"
                lngCol = i * 2 + 3: Exit Function
            End If
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                strErr = "��Сʱ��û������������Сʱ�����ڻ������24Сʱ��"
                lngCol = i * 2 + 3: Exit Function
            End If
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                strErr = "�ķ�����û�����������ķ��������ڻ������60���ӡ�"
                lngCol = i * 2 + 3: Exit Function
            End If
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then
                    strErr = "��ͬһ���ʱ��������ǰһ�ε�ʱ�䡣"
                    lngCol = i * 2 + 3: Exit Function
                End If
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf mstr�����λ = "��" Then
        If mintƵ�ʼ�� = 1 Then
            '8:00-12:00-14:00��8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then
                strErr = "�����˷Ƿ��ַ���": Exit Function
            End If
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> mintƵ�ʴ��� Then
                strErr = "�����˷Ƿ��ַ���": Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then
                    strErr = "�����˶��ʱ��ָ���"":""��"
                    lngCol = i + 2: Exit Function
                End If
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                    strErr = "��Сʱ��û������������Сʱ�����ڻ������24Сʱ��"
                    lngCol = i + 2: Exit Function
                End If
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                    strErr = "�ķ�����û�����������ķ��������ڻ������60���ӡ�"
                    lngCol = i + 2: Exit Function
                End If
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then
                        strErr = "��ʱ��������ǰһ�ε�ʱ�䡣"
                        lngCol = i + 2: Exit Function
                    End If
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00��1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then
                strErr = "�����˷Ƿ��ַ���": Exit Function
            End If
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> mintƵ�ʴ��� Then
                strErr = "�����˷Ƿ��ַ���": Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then
                    strErr = "�����˷Ƿ��ַ���"
                    lngCol = i * 2 + 2: Exit Function
                End If
                
                '�����������
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then
                    strErr = "�������������벻��ȷ��"
                    lngCol = i * 2 + 2: Exit Function
                End If
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > mintƵ�ʼ�� Then
                    strErr = "������������ 1-" & mintƵ�ʼ�� & " ֮�䡣"
                    lngCol = i * 2 + 2: Exit Function
                End If
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then
                        strErr = "����������С����һ�ε�������"
                        lngCol = i * 2 + 2: Exit Function
                    End If
                End If
                
                '����ʱ�䲿��
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then
                    strErr = "�����˶��ʱ��ָ���"":""��"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then
                    strErr = "��Сʱ��û������������Сʱ�����ڻ������24Сʱ��"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                    strErr = "�ķ�����û�����������ķ��������ڻ������60���ӡ�"
                    lngCol = i * 2 + 3: Exit Function
                End If
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then
                        strErr = "��ͬһ���ʱ��������ǰһ�ε�ʱ�䡣"
                        lngCol = i * 2 + 3: Exit Function
                    End If
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf mstr�����λ = "Сʱ" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then
            strErr = "�����˷Ƿ��ַ���": Exit Function
        End If
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> mintƵ�ʴ��� Then
            strErr = "�����˷Ƿ��ַ���": Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then
                strErr = "�����˶��ʱ��ָ���"":""��"
                lngCol = i + 2: Exit Function
            End If
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > mintƵ�ʼ�� Or Split(strTmp, ":")(0) = "" Then
                strErr = "��Сʱ��û������������Сʱ������ 1-" & mintƵ�ʼ�� & "Сʱ ֮�䡣"
                lngCol = i + 2: Exit Function
            End If
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then
                strErr = "�ķ�����û�����������ķ��������ڻ������60���ӡ�"
                lngCol = i + 2: Exit Function
            End If
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then
                    strErr = "��ʱ��������ǰһ�ε�ʱ�䡣"
                    lngCol = i + 2: Exit Function
                End If
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'���ܣ�����ַ����Ƿ�ֻ����ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Private Sub mshTime_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long
    
    If Row = mshTime.MsfObj.FixedRows And mshTime.Rows = mshTime.MsfObj.FixedRows + 1 Then
        Cancel = True
        For i = 1 To mshTime.Cols - 1
            mshTime.TextMatrix(Row, i) = ""
        Next
        Call AdjustOrder(Row)
        mblnChange = True
    End If
End Sub

Private Sub mshTime_EditChange(curText As String)
    If Visible Then mblnChange = True
End Sub

Private Sub mshTime_EditKeyPress(KeyAscii As Integer)
    If mshTime.ColData(mshTime.Col) = 4 Then
        If InStr("01234567890:" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub mshTime_EnterCell(Row As Long, Col As Long)
    If mshTime.ColData(Col) = 4 Then
        mshTime.MaxLength = 5
    Else
        mshTime.MaxLength = 0
    End If
End Sub

Private Sub mshTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strText As String, strLike As String
    Dim vPoint As POINTAPI
    
    If KeyCode = 13 And mshTime.Col = 1 And mshTime.TxtVisible Then
        '�÷���Ϣ
        strLike = gstrMatch
        strText = Replace(UCase(mshTime.Text), "'", "''")
        strSql = _
            " Select Distinct ID,����ID,����,���� From (" & _
            " Select 0 as ����ID,0 as ID,'-' as ����,'<��ȷ��>' as ����,NULL as ����ID,NULL as ����ID From Dual Union ALL" & _
            " Select 1 as ����ID,A.ID,A.����,A.����,B.���� as ����ID,B.���� as ����ID" & _
            " From ������ĿĿ¼ A,������Ŀ���� B" & _
            " Where A.���='E' And A.��������='" & IIf(mint���÷�Χ = 1, 2, 4) & "'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And A.ID=B.������ĿID)" & _
            " Where (���� Like '" & strText & "%'" & _
            " Or Upper(����) Like '" & strLike & strText & "%'" & _
            " Or Upper(����ID) Like '" & strLike & strText & "%'" & _
            " Or Upper(����ID) Like '" & strLike & strText & "%')" & _
            " Order by ����ID,����"
        With mshTime.MsfObj
            vPoint = zlControl.GetCoordPos(.hWnd, .CellLeft - 30, .CellTop - 45)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "�÷�", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        End With
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ���" & IIf(mint���÷�Χ = 1, "��ҩ;��", "��ҩ�÷�") & "��", vbInformation, gstrSysName
            End If
            mshTime.TxtVisible = False
            Cancel = True
        Else
            mshTime.TxtVisible = False
            mshTime.RowData(mshTime.Row) = rsTmp!ID
            mshTime.TextMatrix(mshTime.Row, mshTime.Col) = IIf(rsTmp!���� = "-", rsTmp!����, rsTmp!���� & "-" & rsTmp!����)
            mblnChange = True
        End If
    End If
End Sub

Private Sub mshTime_CommandClick()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    '�÷���Ϣ
    strSql = _
        " Select 0 as ����ID,0 as ID,'-' as ����,'<��ȷ��>' as ���� From Dual Union ALL" & _
        " Select 1 as ����ID,A.ID,A.����,A.���� From ������ĿĿ¼ A" & _
        " Where A.���='E' And A.��������='" & IIf(mint���÷�Χ = 1, 2, 4) & "'" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
        " Order by ����ID,����"
    With mshTime.MsfObj
        vPoint = zlControl.GetCoordPos(.hWnd, .CellLeft - 30, .CellTop - 45)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "�÷�", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
    End With
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û��" & IIf(mint���÷�Χ = 1, "��ҩ;��", "��ҩ�÷�") & "����,���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        End If
    Else
        mshTime.RowData(mshTime.Row) = rsTmp!ID
        mshTime.TextMatrix(mshTime.Row, mshTime.Col) = IIf(rsTmp!���� = "-", rsTmp!����, rsTmp!���� & "-" & rsTmp!����)
    End If
End Sub

Private Sub mshTime_KeyPress(KeyAscii As Integer)
    If mshTime.ColData(mshTime.Col) = 4 Then
        If InStr("01234567890:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
