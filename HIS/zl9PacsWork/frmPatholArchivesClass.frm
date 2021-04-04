VERSION 5.00
Begin VB.Form frmPatholArchivesClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   Icon            =   "frmPatholArchivesClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdSave 
      Caption         =   "�� ��(&S)"
      Height          =   400
      Left            =   10080
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ ��(&D)"
      Height          =   400
      Left            =   8760
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9975
      GridRows        =   501
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
   End
End
Attribute VB_Name = "frmPatholArchivesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function IsAllowDelArchivesClass(ByVal lngClassID As Long) As Boolean
'�Ƿ�����ɾ�������������
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select ����ID from ��������Ϣ where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngClassID)
    
    
    IsAllowDelArchivesClass = IIf(rsData.RecordCount > 0, False, True)
End Function



Private Sub cmdDel_Click()
'��������������ݣ���ʹ�õ�����ܽ���ɾ��
On Error GoTo ErrHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ĵ������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ĵ������", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�ж��Ƿ�����ɾ�������
    If Not IsAllowDelArchivesClass(ufgData.KeyValue(ufgData.SelectionRow)) Then
        Call MsgBoxD(Me, "�õ�������ѱ�ʹ�ã����ܽ���ɾ����", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "ȷ��Ҫɾ��ѡ��ĵ��������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    'ɾ����
    Call ufgData.DelCurRow
    
    '����ɾ���ĵ����������
    Call SaveArchivesClassData(True)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    Dim blnValid As Boolean
    
    '������𱣴�
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "û���ҵ���Ҫ����ĵ��������Ϣ����¼�뵵��������ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽��������б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        
        ufgData.SetFocus
        
        Exit Sub
    End If
    
    Call SaveArchivesClassData
    
    Call MsgBoxD(Me, "�����ѳɹ����档", vbOKOnly, Me.Caption)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub SaveArchivesClassData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'------------------------------------------------------------------------------
'blnIsSaveOnlyDel:�Ƿ��������ɾ��������
'------------------------------------------------------------------------------
'������𱣴�


    Dim i As Long
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim dtSerivcesTime As Date
    
    Dim strNewId As String
    

    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
            
            dtSerivcesTime = zlDatabase.Currentdate
            strSQL = "select Zl_������_��������([1],[2],[3],[4],[5],[6])  as ����ֵ from dual"
            
            Set rsResult = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                    ufgData.Text(i, gstrArchivesClass_��������), _
                                                    Val(ufgData.Text(i, gstrArchivesClass_��������)), _
                                                    ufgData.Text(i, gstrArchivesClass_��������), _
                                                    UserInfo.����, _
                                                    CDate(dtSerivcesTime), _
                                                    ufgData.Text(i, gstrArchivesClass_��ע) _
                                                    )
                                                            
            
            If rsResult.RecordCount <= 0 Then
                Call err.Raise(0, "SaveArchivesData", "δ�ɹ���ȡ������ĵ�������ID,����ʧ�ܡ�")
                Exit Sub
            End If
            
            '���µ��������б�
            ufgData.Text(i, gstrArchivesClass_ID) = Nvl(rsResult!����ֵ)
            ufgData.Text(i, gstrArchivesClass_������) = UserInfo.����
            ufgData.Text(i, gstrArchivesClass_����ʱ��) = dtSerivcesTime
            
        ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
            
            strSQL = "Zl_������_���·���('" & ufgData.KeyValue(i) & "','" & _
                                                ufgData.Text(i, gstrArchivesClass_��������) & "'," & _
                                                Val(ufgData.Text(i, gstrArchivesClass_��������)) & ",'" & _
                                                ufgData.Text(i, gstrArchivesClass_��������) & "','" & _
                                                ufgData.Text(i, gstrArchivesClass_��ע) & "')"
            
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
        ElseIf ufgData.RowState(i) = TDataRowState.Del Then
            'ɾ�����������¼
            If Trim(ufgData.KeyValue(i)) <> "" Then
                strSQL = "Zl_������_ɾ������('" & ufgData.KeyValue(i) & "')"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
        
        
        '������״̬
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
    
End Sub



Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitArchivesClassList
    
    Call LoadArchivesClassData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitArchivesClassList()
'���õ���������ʾ�б�

    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
    ufgData.DefaultColNames = gstrArchivesClassCols
    ufgData.ColNames = gstrArchivesClassCols
    ufgData.ColConvertFormat = gstrArchivesClassConvertFormat
End Sub



Private Sub LoadArchivesClassData()
'���뵵����������
    Dim strSQL As String
    
    strSQL = "select ID, ��������,��������,��������, ��ע,������,����ʱ�� from ���������� order by ����ʱ��"
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim strNewArchivesClassName
    Dim iCol As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        Exit Sub
    End If
    
    
    If Col = ufgData.GetColIndex(gstrArchivesClass_��������) Then
    '��鵵������Ƿ��ظ�
    
        strNewArchivesClassName = ufgData.CheckEquateValue(Row, Col)
        If strNewArchivesClassName <> "" Then
            Call MsgBoxD(Me, "������� [" & ufgData.Text(Row, gstrArchivesClass_��������) & "]�Ѿ����ڡ�", vbOKOnly, Me.Caption)
            
            ufgData.Text(Row, gstrArchivesClass_��������) = strNewArchivesClassName
        End If
    End If
    
    
    '���δ¼��������ƣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrArchivesClass_��������)
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrArchivesClass_��������) = "", ufgData.ErrCellColor, ufgData.BackColor)
          
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
