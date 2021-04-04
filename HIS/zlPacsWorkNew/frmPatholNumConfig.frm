VERSION 5.00
Begin VB.Form frmPatholNumConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����ű�����"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��"
      Height          =   400
      Left            =   7680
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   400
      Left            =   6240
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�� ��(&S)"
      Height          =   400
      Left            =   9120
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   10560
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.OptionButton optYear 
      Caption         =   "4λ"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   5730
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optMonth 
      Caption         =   "2λ"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   5730
      Width           =   615
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8705
      DefaultCols     =   ""
      GridRows        =   11
      HeadCheckValue  =   1
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      GridLineColor   =   14737632
      ExtendLastCol   =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "������ʹ��              ���"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   2535
   End
End
Attribute VB_Name = "frmPatholNumConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_str��ʾ = "��ʾ"

Private Function reInitData() As Boolean
    Dim strSql As String
    On Error GoTo errH
    reInitData = False
    gcnOracle.BeginTrans
    
    strSql = "ZL_����������_Insert(1,0,'CG',1,1,1,3,4,1,'����')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_����������_Insert(2,1,'BD',1,1,1,3,4,1,'����')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_����������_Insert(3,2,'XB',1,1,1,3,4,1,'ϸ��')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_����������_Insert(4,3,'HZ',1,1,1,3,4,1,'����')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_����������_Insert(5,4,'SJ',1,1,1,3,4,1,'ʬ��')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_����������_Insert(6,5,'KSSL',1,1,1,3,4,1,'����ʯ��')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    gcnOracle.CommitTrans
    reInitData = True
    Exit Function
errH:
    Call gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Sub LordPatholNumRules()
'���벡��Ź�����ʾ���б���
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    Dim curDate As Date
    Dim i As Integer
    Dim intID As Integer
    
    On Error GoTo errH
    
    strSql = "select ID,����,ǰ׺,��,��,��,���λ��,���λ��,��ʼ��,����  from ����������  order by ID"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set ufgData.AdoData = rsData
    
    If rsData.RecordCount > 5 Then '���������������
    
    ElseIf rsData.RecordCount = 0 Then '�����������Ϊ0����Ҫ��ʼ��
        If reInitData() = False Then
            Call MsgBoxD(Me, "�������������ݳ�ʼ��ʧ�ܣ�����ϵ���ά����Ա���", vbOKOnly, "�����������")
            Exit Sub
        End If
    Else '
        Call MsgBoxD(Me, "���������������쳣������ϵ���ά����Ա���", vbOKOnly, "�����������")
        Exit Sub
    End If

    Call ufgData.RefreshData
    
    curDate = zlDatabase.Currentdate
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.Text(i, gstrPatholNumSet_ID) <> "" Then

            strSql = "select ��ǰ��� as ��ʼ�� from ��������¼ where �������ID=[1]"
            strWhere = ""
            
            If InStr(ufgData.Text(i, gstrPatholNumSet_��), "��") > 0 Then
                strWhere = strWhere & " and ��=[2]"
            End If

            If InStr(ufgData.Text(i, gstrPatholNumSet_��), "��") > 0 Then
                strWhere = strWhere & " and ��=[3]"
            End If

            If InStr(ufgData.Text(i, gstrPatholNumSet_��), "��") > 0 Then
                strWhere = strWhere & " and ��=[4]"
            End If
 
            Set rsData = zlDatabase.OpenSQLRecord(strSql & strWhere, Me.Caption, ufgData.Text(i, gstrPatholNumSet_ID), Val(Format(curDate, "yyyy")), Val(Format(curDate, "mm")), Val(Format(curDate, "dd")))
            
            If rsData.RecordCount = 0 Then
                intID = Val(ufgData.Text(i, gstrPatholNumSet_ID)) - 1
                strSql = "select ��ǰ��� as ��ʼ�� from ��������¼ where ����=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSql & strWhere, Me.Caption, intID, Val(Format(curDate, "yyyy")), Val(Format(curDate, "mm")), Val(Format(curDate, "dd")))
                If rsData.RecordCount > 0 Then
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = Nvl(rsData!��ʼ��, "0")
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = Val(ufgData.Text(i, gstrPatholNumSet_��ʼ��)) + 1
                Else
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = 1
                End If
            
            Else
                If rsData.RecordCount > 0 Then
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = Nvl(rsData!��ʼ��, "0")
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = Val(ufgData.Text(i, gstrPatholNumSet_��ʼ��)) + 1
                Else
                    ufgData.Text(i, gstrPatholNumSet_��ʼ��) = 1
                End If
            End If

        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub initUfgList()
'��ʼ���б�
    
    On Error GoTo errH
    
    ufgData.IsKeepRows = False
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.DefaultColNames = gstrPatholNumSetCols
    ufgData.ColNames = gstrPatholNumSetCols

    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
    '��ֹ��������Ҽ��˵�
    ufgData.IsShowPopupMenu = False
    
    ufgData.ColConvertFormat = gstrPatholNumSetConvertFormat
    
    Exit Sub
errH:
    Call err.Raise(0, , "��ʼ���б�ʧ��")
End Sub

Private Sub cmdAdd_Click()
    Dim lngNewRow As Long
    
    On Error GoTo errH

    lngNewRow = ufgData.NewRow
    
    ufgData.Text(lngNewRow, gstrPatholNumSet_ID) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_����) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_ǰ׺) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_��) = "1-��"
    ufgData.Text(lngNewRow, gstrPatholNumSet_��) = "1-��"
    ufgData.Text(lngNewRow, gstrPatholNumSet_��) = "1-��"
    ufgData.Text(lngNewRow, gstrPatholNumSet_���λ��) = IIf(optYear.value = True, "4", "2")
    ufgData.Text(lngNewRow, gstrPatholNumSet_���λ��) = 4
    ufgData.Text(lngNewRow, gstrPatholNumSet_����) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_��ʼ��) = "1"
    
    Call ufgData_OnAfterEdit(lngNewRow, ufgData.DataGrid.Col)
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str��ʾ)
End Sub

Private Function CheckDelable(ByVal lngID As Long) As Boolean
'����Ƿ��й�����ʹ�ñ����ݣ������ж��Ƿ����ɾ��
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckDelable = False
    
    strSql = "select �������ID from  ��������Ϣ where �������ID=[1] and rownum <2 "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������Ƿ�ʹ��", lngID)
    
    If rsData.RecordCount = 0 Then CheckDelable = True
        
End Function

Private Sub DelRow()
'ɾ����
    Dim lngRow As Long
    Dim strSql As String
    Dim intTMP As Integer
    
    On Error GoTo errH
    
    lngRow = ufgData.SelectionRow
    
    '����ID �ж��ܷ�ɾ����û��ID˵���Ǹ��½��ģ���ʾȷ�Ϻ�ɾ��
    If ufgData.Text(lngRow, gstrPatholNumSet_ID) <> "" Then
    
        intTMP = Val(ufgData.Text(lngRow, gstrPatholNumSet_ID))
        
        If intTMP >= 1 And intTMP <= 6 Then
            Call MsgBoxD(Me, "�ú���������ڻ�������,����ɾ��", vbOKOnly, C_str��ʾ)
            Exit Sub
        Else
            If CheckDelable(Val(ufgData.Text(lngRow, gstrPatholNumSet_ID))) = True Then
            
                If MsgBoxD(Me, "�����濼���Ƿ�ɾ���ú������?", vbYesNo + vbDefaultButton2 + vbCritical, C_str��ʾ) = vbNo Then Exit Sub
                
                strSql = "ZL_����������_Delete(" & ufgData.Text(lngRow, gstrPatholNumSet_ID) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                Call ufgData.DelRow(lngRow, False)
                  
            Else
                Call MsgBoxD(Me, "�ú����������ʹ��,����ɾ��", vbOKOnly, C_str��ʾ)
            End If
        End If
    Else
        If MsgBoxD(Me, "�Ƿ�ȷ��ɾ���ú������", vbYesNo + vbDefaultButton2, C_str��ʾ) = vbNo Then Exit Sub
        
        Call ufgData.DelRow(lngRow, False)
        Call ufgData.RefreshData
    End If
    
    
    Exit Sub
errH:
    Call err.Raise(0, , "ɾ������ʧ��" & err.Description)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errH
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ������Ŀ,", vbOKOnly, C_str��ʾ)
        Exit Sub
    End If

    Call DelRow
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str��ʾ)
End Sub

Private Function CheckHaveErrCell() As Boolean
'�ж��Ƿ������ݴ��󣨸�����ɫ��
    Dim i As Integer
    Dim j As Integer
    Dim iCol As Integer
    
    On Error GoTo errH
    
    CheckHaveErrCell = True

    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            For j = 1 To ufgData.GridCols - 1
                If ufgData.CellColor(i, j) = ufgData.ErrCellColor Then
                    CheckHaveErrCell = False
                    Exit Function
                End If
            Next
        End If
    Next
    Exit Function
errH:
    Call err.Raise(0, , "�ж��Ƿ�����Ч����ʧ��" & err.Description)
End Function

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim iCol As Integer
    Dim intMax As Integer
    Dim intID As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    
    If CheckHaveErrCell = False Then
        Call MsgBoxD(Me, "������Ч���ݣ���ֹ���棬��ע���޸ĺ�ɫ����", vbOKOnly, C_str��ʾ)
        Exit Sub
    End If

    Call gcnOracle.BeginTrans
    For i = 1 To ufgData.GridRows - 1
    
        If ufgData.RowState(i) <> TDataRowState.Del Then 'ɾ��״̬�Ĳ����棬�����ʱɾ����Ч
        
            strSql = " select max(ID) as ���ID from ���������� "
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ���������ID")
            
            intMax = rsData!���ID + 1
            
            If intMax > 90 And intMax < 100 Then
                Call MsgBoxD(Me, "�������������������������ƣ��뾡��֪ͨ���ݿ������Ա����", vbOKOnly, C_str��ʾ)
            ElseIf intMax > 99 Then
                Call MsgBoxD(Me, "���������������Ѿ��������ƣ��޷����������������뾡��֪ͨ���ݿ������Ա����", vbOKOnly, C_str��ʾ)
                Exit Sub
            End If
            
    
            If ufgData.Text(i, gstrPatholNumSet_ID) = "" Then
                intID = intMax
            Else
                intID = Val(ufgData.Text(i, gstrPatholNumSet_ID))
            End If

                                                            
            strSql = "ZL_����������_Insert(" & intID & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_����)) & ",'" & _
                                                            ufgData.Text(i, gstrPatholNumSet_ǰ׺) & "'," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_��)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_��)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_��)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_���λ��)) & "," & _
                                                            IIf(optYear.value = True, 4, 2) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_��ʼ��)) & ",'" & _
                                                            ufgData.Text(i, gstrPatholNumSet_����) & "')"
                                                            
                                                            
            Call zlDatabase.ExecuteProcedure(strSql, "����������_�½�")
            ufgData.Text(i, gstrPatholNumSet_ID) = intID
            
        End If
    Next
    
    Call gcnOracle.CommitTrans
    Me.Hide

    Exit Sub
errH:
    Call gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Call initUfgList
    Call LordPatholNumRules
    
    Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    Dim i As Integer
    Dim iCol As Integer
    Dim blEx As Boolean  '�����ж�ǰ׺
    Dim intTMP As Integer
    Dim lngSelectRow As Long 'ѡ�е�����
    Dim strTMPEx As String
    Dim strTMPName As String
    
    lngSelectRow = ufgData.SelectionRow
    
    If ufgData.GridRows < 2 Then Exit Sub
    
    blEx = False
    
    With ufgData
       '�Ƚ���ɫ�ָ�Ϊ����
       For i = 1 To .GridRows - 1
           
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_����)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_ǰ׺)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_����)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_��ʼ��)) = .BackColor
       Next
       
       '���ɱ༭����ɫDisCellColor
       For i = 1 To .GridRows - 1
           intTMP = Val(.Text(i, gstrPatholNumSet_ID))
           If intTMP >= 1 And intTMP <= 6 Then
               .CellColor(i, .GetColIndex(gstrPatholNumSet_����)) = .DisCellColor
           End If
       Next
       
      ' �ж�ǰ׺�����ơ������Ƿ�Ϊ��
       For i = 1 To .GridRows - 1
           
           iCol = .GetColIndex(gstrPatholNumSet_����)
           If Trim(.Text(i, gstrPatholNumSet_����)) = "" Then
               .CellColor(i, iCol) = .ErrCellColor
           End If
           
           iCol = .GetColIndex(gstrPatholNumSet_����)
           If Trim(.Text(i, gstrPatholNumSet_����)) = "" Then
               .CellColor(i, iCol) = .ErrCellColor
           End If
       Next
       
       '�ж�ǰ׺�������Ƿ��ظ�
       For i = 1 To .GridRows - 1
           strTMPName = .Text(lngSelectRow, gstrPatholNumSet_����)
           
           If (i <> lngSelectRow) And (.Text(i, gstrPatholNumSet_����) = strTMPName And Trim(strTMPName) <> "") Then
               iCol = .GetColIndex(gstrPatholNumSet_����)
               .CellColor(lngSelectRow, iCol) = .ErrCellColor
           End If
       Next
       
       '��ʼ��
       iCol = .GetColIndex(gstrPatholNumSet_��ʼ��)
       
       For i = 1 To .GridRows - 1
           If Val(.Text(i, gstrPatholNumSet_��ʼ��)) < 0 Then
               .CellColor(i, iCol) = .ErrCellColor
               Call MsgBoxD(Me, "��ʼ��ֻ��Ϊ��С��0������,����", vbOKOnly, C_str��ʾ)
               Exit Sub
           End If
       Next
    
       
       '�ж�ǰ׺���ȳ���5
       iCol = .GetColIndex(gstrPatholNumSet_ǰ׺)
       
       If .CellColor(Row, iCol) <> .ErrCellColor Then
       
           For i = 1 To Len(.Text(Row, gstrPatholNumSet_ǰ׺))
              
               intTMP = Asc(Mid(.Text(Row, gstrPatholNumSet_ǰ׺), i, 1))
               
               '��һ����������ĸ��������ʾΪ��ɫ
               If intTMP <= 47 Or (intTMP >= 58 And intTMP <= 64) Or (intTMP >= 91 And intTMP <= 96) Or intTMP >= 123 Then
                   .CellColor(Row, iCol) = .ErrCellColor
                   blEx = True
                   Exit For
               End If
           Next
           
           If Len(.Text(Row, gstrPatholNumSet_ǰ׺)) > 5 Then
               .CellColor(Row, iCol) = .ErrCellColor
               blEx = True
           End If
       End If
       
    End With
    
    If blEx = True Then Call MsgBoxD(Me, "ǰ׺�ַ������Ϊ5,��ֻ�������ֻ���ĸ,����", vbOKOnly, C_str��ʾ)
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str��ʾ)
End Sub

Private Sub ufgData_OnKeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
' ��ֹ����һЩ����",:'"֮��ķ���
    If (KeyAscii >= 37 And KeyAscii <= 43) Or (KeyAscii >= 58 And KeyAscii <= 63) Or KeyAscii = 44 Or KeyAscii = 46 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errH
    If 1 <= Val(ufgData.KeyValue(ufgData.SelectionRow)) And Val(ufgData.KeyValue(ufgData.SelectionRow)) <= 6 Then
        '��ͼ�༭����������ƣ����ֹ�༭
        If Col = ufgData.GetColIndex(gstrPatholNumSet_����) Then
            Cancel = True
            Call MsgBoxD(Me, "�����������͵����Ʋ����޸�,", vbOKOnly, C_str��ʾ)
            Exit Sub
        End If
    End If
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str��ʾ)
End Sub

