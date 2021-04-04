VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmChargeGroupItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�շ���Ŀ���"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmChargeGroupItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCanCel 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   7680
      TabIndex        =   2
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   300
      Left            =   180
      TabIndex        =   3
      Top             =   4470
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit mshGroupItem 
      Height          =   3510
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   6191
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin VB.Label lblEdit 
      Caption         =   "    �û������ѵķ�ʽ���ӵ��շ���Ŀ��ɣ���Ҫ�����û����ڴ�ӡ������"
      Height          =   435
      Index           =   13
      Left            =   990
      TabIndex        =   4
      Top             =   210
      Width           =   6810
   End
   Begin VB.Image img���� 
      Height          =   600
      Left            =   120
      Picture         =   "frmChargeGroupItem.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmChargeGroupItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemGroupID As Long             '�շ���ĿID
Sub Init()
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    With mshGroupItem
        .Cols = 6
        .ColWidth(0) = 1500
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 4000
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "�۸�"
        .TextMatrix(0, 2) = "���"
        .TextMatrix(0, 3) = "���㵥λ"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "˵��"
        .ColAlignment(2) = 1
        'ʵ�ַ�ʽ
        .ColData(0) = 1 '��ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        .ColData(1) = 4
        .ColData(2) = 4
        .ColData(3) = 4
        .ColData(4) = 4
        .ColData(5) = 4
        
        .PrimaryCol = 1
        .Active = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Function ShowMe(objfrm As Object, ItemID As Long)
    '''''''''''''''''''''''''''''''''''''''''
    '����               ���ϼ��������
    '����               objfrm �ϼ��������
    '''''''''''''''''''''''''''''''''''''''''
    ItemGroupID = ItemID
    Me.Show vbModal, objfrm
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    Call SaveDate
    Unload Me
End Sub
Private Sub Form_Load()
    Call GetPriceGrade(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    Call Init
    Call LoadDate
End Sub

Private Function IsValid() As Boolean
    Dim strTmp As String
    '��������Ƿ�Ϸ�
    With Me.mshGroupItem
        For i = 1 To .Rows - 1
            If .Row <> .Rows - 1 Then
                If Len(Trim(.TextMatrix(i, 0))) <= 0 Then
                    MsgBox "��" & i & "�е����Ʋ���Ϊ��!", vbQuestion, gstrSysName
                    .Row = i
                    .Col = 0
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If zlCommFun.StrIsValid(.TextMatrix(i, 0), 80) = False Then
                .Row = i
                .Col = 0
                .SetFocus
                .TxtSetFocus
                Exit Function
            End If
            If .TextMatrix(i, 1) <> "" Then
                If IsNumeric(.TextMatrix(i, 1)) = False Then
                    MsgBox "��" & i & "�еļ۸����벻��ȷ!", vbQuestion, gstrSysName
                    .Row = i
                    .Col = 1
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If .TextMatrix(i, 2) <> "" Then
                If zlCommFun.StrIsValid(.TextMatrix(i, 2), 40) = False Then
                    .Row = i
                    .Col = 2
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If .TextMatrix(i, 3) <> "" Then
                If zlCommFun.StrIsValid(.TextMatrix(i, 3), 40) = False Then
                    .Row = i
                    .Col = 3
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If .TextMatrix(i, 4) <> "" Then
                If IsNumeric(.TextMatrix(i, 4)) = False Then
                    MsgBox "��" & i & "�е��������벻��ȷ!", vbQuestion, gstrSysName
                    .Row = i
                    .Col = 4
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If .TextMatrix(i, 5) <> "" Then
                If zlCommFun.StrIsValid(.TextMatrix(i, 5), 500) = False Then
                    .Row = i
                    .Col = 5
                    .SetFocus
                    .TxtSetFocus
                    Exit Function
                End If
            End If
            If InStr(1, strTmp & ",", "," & .TextMatrix(i, 0) & ",") > 0 Then
                MsgBox "��" & i & "�����Ƴ����ظ�!", vbQuestion, gstrSysName
                .Row = i
                .Col = 0
                .SetFocus
                .TxtSetFocus
                Exit Function
            End If
            strTmp = strTmp & "," & .TextMatrix(i, 0)
        Next
    End With
    IsValid = True
    
End Function

Private Function Check�ظ���Ŀ(strName As String) As Boolean
    Dim n As Integer
    
    Check�ظ���Ŀ = True
    
    With mshGroupItem
        If .Rows < 1 Then Exit Function
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 0) = strName Then
                Check�ظ���Ŀ = False
                Exit Function
            End If
        Next
    End With
    
End Function

Sub SaveDate()
    Dim lngId As Long
    On Error GoTo errH
    gstrSQL = "ZL_�շ���Ŀ���_DELETE(" & ItemGroupID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    With Me.mshGroupItem
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) > 0 Then
                lngId = sys.NextId("�շ���Ŀ���")
                gstrSQL = "ZL_�շ���Ŀ���_INSERT(" & lngId & "," & ItemGroupID & ",'" & _
                          .TextMatrix(i, 0) & "'," & Val(IIF(IsNull(.TextMatrix(i, 1)), 0, .TextMatrix(i, 1))) & _
                          " ,'" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "'," & _
                          Val(IIF(IsNull(.TextMatrix(i, 4)), 0, .TextMatrix(i, 4))) & ",'" & _
                          .TextMatrix(i, 5) & "')"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub
Sub LoadDate()
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ����,����,���,���㵥λ,����,˵�� from �շ���Ŀ��� where �շ���Ŀid = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, ItemGroupID)
    
    With Me.mshGroupItem
        .Row = 1
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, 0) = rsTmp("����")
            .TextMatrix(.Rows - 1, 1) = IIF(rsTmp("����") = 0, "", rsTmp("����"))
            .TextMatrix(.Rows - 1, 2) = IIF(IsNull(rsTmp("���")), "", rsTmp("���"))
            .TextMatrix(.Rows - 1, 3) = IIF(IsNull(rsTmp("���㵥λ")), "", rsTmp("���㵥λ"))
            .TextMatrix(.Rows - 1, 4) = IIF(rsTmp("����") = 0, "", rsTmp("����"))
            .TextMatrix(.Rows - 1, 5) = IIF(IsNull(rsTmp("˵��")), "", rsTmp("˵��"))
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshGroupItem_CommandClick()
    Dim strReturn As String
    Dim rsTmp As New ADODB.Recordset
    Dim n As Integer
    Dim strWherePriceGrade As String
    
    On Error GoTo ErrHandle
    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [1])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [2])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And d.�۸�ȼ� = [3])" & vbNewLine & _
            "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [1])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [3])))))"
    End If

    With mshGroupItem
        gstrSQL = _
            "SELECT A.����,A.����,A.���,A.���㵥λ," & _
            " ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.00'))) �۸�,A.ID" & _
            " FROM �շ���ĿĿ¼ A,�շѼ�Ŀ D" & _
            " WHERE A.ID=D.�շ�ϸĿID(+) And a.ID>0 And A.��� Not In ('5', '6', '7') " & _
            "       And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)" & _
            "       And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)" & _
                    strWherePriceGrade & vbNewLine & _
            " Group By A.����,A.����,A.���,A.���㵥λ,A.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        If rsTmp.RecordCount < 1 Then Exit Sub
        
        strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1000,0,2;����,2500,0,2;���,1500,0,2;��λ,1500,0,2;�۸�,1000,1,2;ID,0,0,2", _
            "��Ŀѡ����", True, strTemp, 3, 1000 + 2500 + 1500 + 1500 + 1000 + 1800)
        If Trim(strReturn) = "" Then Exit Sub
        
        If Check�ظ���Ŀ(CStr(Split(strReturn, ",")(1))) = False Then
            MsgBox "�Ѿ�������Ϊ[" & Split(strReturn, ",")(1) & " ]�������Ŀ��", vbQuestion, gstrSysName
            Exit Sub
        End If
        
        '����,�۸�,���,���㵥λ,����,˵��
        .TextMatrix(.Row, 0) = Split(strReturn, ",")(1)
        .TextMatrix(.Row, 1) = FormatEx(Split(strReturn, ",")(4), 2)
        .TextMatrix(.Row, 2) = Split(strReturn, ",")(2)
        .TextMatrix(.Row, 3) = Split(strReturn, ",")(3)
        .TextMatrix(.Row, 4) = "1"
        .TextMatrix(.Row, 5) = ""
         
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub mshGroupItem_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsTmp As New ADODB.Recordset
    Dim strWherePriceGrade As String
    
    On Error GoTo ErrHandle
    If KeyCode = 13 Then
        With Me.mshGroupItem
            If .Col = 0 Then
                If .Text = "" Then
                    If .Row = .Rows - 1 Then
                        cmdOK.SetFocus
                    End If
                    Exit Sub
                End If
                
                .Text = UCase(Trim(.Text))
                strKey = .Text
                If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
                    strWherePriceGrade = " And d.�۸�ȼ� Is Null"
                Else
                    strWherePriceGrade = "" & _
                        " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [3])" & vbNewLine & _
                        "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [4])" & vbNewLine & _
                        "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And d.�۸�ȼ� = [5])" & vbNewLine & _
                        "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
                        "          And Not Exists (Select 1" & vbNewLine & _
                        "                          From �շѼ�Ŀ" & vbNewLine & _
                        "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                        "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
                        "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [4])" & vbNewLine & _
                        "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [5])))))"
                End If
                gstrSQL = _
                    "SELECT A.����,A.����,A.���,A.���㵥λ," & _
                    " ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.00'))) �۸�,A.ID" & _
                    " FROM �շ���ĿĿ¼ A,�շѼ�Ŀ D" & _
                    " WHERE A.ID=D.�շ�ϸĿID(+) And a.ID>0 And A.��� Not In ('5', '6', '7') " & _
                    "       And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)" & _
                    "       And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)" & _
                            strWherePriceGrade & vbNewLine & _
                    "       And (A.���� Like [1] Or A.���� Like [2] " & _
                    "           Or Exists (Select 1 From �շ���Ŀ���� Where �շ�ϸĿid = a.Id And ���� Like [2])) " & _
                    " Group By A.����,A.����,A.���,A.���㵥λ,A.ID"
                    
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey & "%", "%" & strKey & "%", _
                    gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
                
                If rsTmp.RecordCount < 1 Then
                    .TextMatrix(.Row, 0) = .Text
                    Exit Sub
                End If
                
                If rsTmp.RecordCount > 1 Then
                    strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1000,0,2;����,2500,0,2;���,1500,0,2;��λ,1500,0,2;�۸�,1000,1,2;ID,0,0,2", _
                        "��Ŀѡ����", True, strTemp, 3, 1000 + 2500 + 1500 + 1500 + 1000 + 1800)
                    If Trim(strReturn) = "" Then Exit Sub
                    
                    If Check�ظ���Ŀ(CStr(Split(strReturn, ",")(1))) = False Then
                        MsgBox "�Ѿ�������Ϊ[" & Split(strReturn, ",")(1) & " ]�������Ŀ��", vbQuestion, gstrSysName
                        Exit Sub
                    End If
                    
                    '����,�۸�,���,���㵥λ,����,˵��
                    .Text = Split(strReturn, ",")(1)
                    .TextMatrix(.Row, 0) = .Text
                    .TextMatrix(.Row, 1) = FormatEx(Split(strReturn, ",")(4), 2)
                    .TextMatrix(.Row, 2) = Split(strReturn, ",")(2)
                    .TextMatrix(.Row, 3) = Split(strReturn, ",")(3)
                Else
                    '����,�۸�,���,���㵥λ,����,˵��
                    .Text = rsTmp!����
                    .TextMatrix(.Row, 0) = .Text
                    .TextMatrix(.Row, 1) = FormatEx(rsTmp!�۸�, 2)
                    .TextMatrix(.Row, 2) = nvl(rsTmp!���, "")
                    .TextMatrix(.Row, 3) = nvl(rsTmp!���㵥λ, "")
                End If
                .TextMatrix(.Row, 4) = "1"
                .TextMatrix(.Row, 5) = ""
            End If
            
            If .TextMatrix(.Row, 5) = "" And .Col = 5 And .Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
            End If
            If .TextMatrix(.Row, 5) = "" And .Col = 5 And .Row < .Rows - 1 Then
                .TextMatrix(.Row, 5) = "  "
            End If
            If .Col > 0 And .Col < 5 And .TextMatrix(.Row, .Col) = "" Then
                If .Text = "" Then
                    OS.PressKey vbKeyRight
                End If
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
