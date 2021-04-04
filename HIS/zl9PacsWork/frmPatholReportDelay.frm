VERSION 5.00
Begin VB.Form frmPatholReportDelay 
   Caption         =   "�����ӳ�"
   ClientHeight    =   7095
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10875
   Icon            =   "frmPatholReportDelay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10875
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   10335
      TabIndex        =   5
      Top             =   6000
      Width           =   10335
      Begin VB.CommandButton cmdPrint 
         Caption         =   "�� ӡ(&P)"
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "���ݱ���(&S)"
         Height          =   400
         Left            =   9000
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ɾ����¼(&C)"
         Height          =   400
         Left            =   7680
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame framReportDelay 
      Caption         =   "�����ӳټ�¼"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10335
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         _ExtentX        =   17171
         _ExtentY        =   4471
         GridRows        =   21
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholReportDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngCurAdviceId As Long
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf

Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigReportDelayFace(False, "ҽ��ID��Ч���顣")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId


    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)

    If Trim(mrecStudyInf.strPatholNumber) = "" Then
        Call ConfigReportDelayFace(False, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "�ü����δ������Ч�Ĳ���ţ���ȷ�ϸü���Ƿ��ѱ����ա�", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigReportDelayFace(True)
    End If

    Call LoadReportDelayData(mrecStudyInf.lngPatholAdviceId)
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub



Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'����Ȩ��
    Dim blnIsAllowDelay As Boolean
    
    blnIsAllowDelay = CheckPopedom(mstrPrivs, "�����ӳ�")
    
    cmdCancel.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    cmdSave.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    
    cmdPrint.Enabled = blnIsAllowDelay
    
    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub ConfigReportDelayFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'�����ؼ����
    cmdSave.Enabled = blnIsValid
    cmdCancel.Enabled = blnIsValid
    cmdPrint.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub AdjustFace()
'�������沼��
    framReportDelay.Left = 120
    framReportDelay.Top = 120
    framReportDelay.Width = Me.Width - 360
    framReportDelay.Height = Me.Height - picControl.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framReportDelay.Width - 240
    ufgData.Height = framReportDelay.Height - 360
    
    
    picControl.Left = 120
    picControl.Top = Me.Height - picControl.Height - 620
    picControl.Width = Me.Width - 240
    
    
    cmdPrint.Left = 0
    cmdPrint.Top = 0
    
    cmdSave.Left = picControl.Width - cmdSave.Width - 120
    cmdSave.Top = 0
    
    cmdCancel.Left = cmdSave.Left - cmdCancel.Width - 120
    cmdCancel.Top = 0
End Sub



Private Sub InitReportDelayList()
'��ʼ�������ӳ���ʾ�б�
    Dim strTemp As String
    
    

     '�ж����ݿ�������Ƿ������� �����ȡ���ݿ����  û�������Ĭ��
    strTemp = zlDatabase.GetPara("�����ӳ��б�����", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrReportDelayCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrReportDelayCols
    Else
        ufgData.ColNames = strTemp
    End If
    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrReportDelayConvertFormat
End Sub


Private Sub ufgData_OnColFormartChange()
 '�رմ���ʱ�����б�����
    zlDatabase.SetPara "�����ӳ��б�����", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadReportDelayData(ByVal lngPatholAdviceId As Long)
'���뱨���ӳ�����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID,�ӳ�ԭ��,�ӳ�����,��ʱ���,ת����,�Ǽ���,�Ǽ�ʱ��,��ǰ״̬ from �������ӳ� where ����ҽ��ID=[1] order by �Ǽ�ʱ��"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub SaveReportDelayData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'blnIsSaveOnlyDel:�Ƿ񱣴��ɾ��������

'���汨���ӳ�����
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    For i = 1 To ufgData.GridRows - 1
        Select Case ufgData.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                dtServicesTime = zlDatabase.Currentdate
                
                strSql = "select Zl_�������ӳ�_����([1],[2],[3],[4],[5],[6],[7]) as ����ֵ from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                        mrecStudyInf.lngPatholAdviceId, _
                                                        ufgData.Text(i, gstrReportDelay_�ӳ�ԭ��), _
                                                        Val(ufgData.Text(i, gstrReportDelay_�ӳ�����)), _
                                                        ufgData.Text(i, gstrReportDelay_��ʱ���), _
                                                        ufgData.Text(i, gstrReportDelay_ת����), _
                                                        UserInfo.����, _
                                                        CDate(dtServicesTime))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveReportDelayData", "δ�ɹ���ȡ������ı����ӳ�ID,����ʧ�ܡ�")
                    Exit Sub
                End If
                
                
                ufgData.Text(i, gstrReportDelay_ID) = rsData!����ֵ
                ufgData.Text(i, gstrReportDelay_�Ǽ���) = UserInfo.����
                ufgData.Text(i, gstrReportDelay_��ǰ״̬) = "δ��ӡ"
                ufgData.Text(i, gstrReportDelay_�Ǽ�ʱ��) = dtServicesTime
                                                        
            Case TDataRowState.Del
                strSql = "Zl_�������ӳ�_ɾ��(" & Val(ufgData.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                strSql = "Zl_�������ӳ�_����(" & Val(ufgData.KeyValue(i)) & ",'" & _
                                                ufgData.Text(i, gstrReportDelay_�ӳ�ԭ��) & "'," & _
                                                Val(ufgData.Text(i, gstrReportDelay_�ӳ�����)) & ",'" & _
                                                ufgData.Text(i, gstrReportDelay_��ʱ���) & "','" & _
                                                ufgData.Text(i, gstrReportDelay_ת����) & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

        End Select
        
        '������״̬
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
End Sub


Private Sub cmdCancel_Click()
'ɾ���ײ�
On Error GoTo errHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ɾ���ı����ӳټ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫɾ���ñ����ӳ�������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgData.DelCurRow
    
    '����ɾ��������
    Call SaveReportDelayData(True)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub PrintReportDelay(ByVal lngReportDelayId As Long)
'��ӡ�����ӳٵ�
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_07", Me, "�����ӳ�ID=" & lngReportDelayId, Decode((Val(zlDatabase.GetPara("�Ƿ�ֱ�Ӵ�ӡ", glngSys, glngModul, 0)) = 1), 0, 0, 2))
End Sub


Private Sub cmdPrint_Click()
'�����ӳٵ���ӡ
On Error GoTo errHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���д�ӡ�ı����ӳټ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���д�ӡ�ı����ӳټ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If MsgBoxD(Me, "�ӳٱ�����δ���棬���ܽ��д�ӡ����Ҫ�Զ�������", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        Else
            '���汨���ӳ���Ϣ
            Call SaveReportDelayData
        End If
    End If
    
    Call PrintReportDelay(ufgData.KeyValue(ufgData.SelectionRow))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub CmdRefresh_Click()
''�����ݻָ�����ʼ״̬
'On Error GoTo errHandle
'    Call mvfgReportDelay.RestoreList
'
'    Call mvfgReportDelay.RefreshReadColColor
'
'    Call RefreshRecordInf
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "��⵽�����ӳ��б��д�����Ч���ݣ���ȷ����������Ƿ���ȷ������¼�룬����ɫ����ǵĵ�Ԫ��Ϊ��¼���ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '�����ײ���Ϣ
    Call SaveReportDelayData
    
    Call MsgBoxD(Me, "�����ѱ���ɹ���", vbOKOnly, Me.Caption)
'    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitReportDelayList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set zlReport = Nothing
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
        
    
    '���δ¼��걾���ƣ�����ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrReportDelay_�ӳ�ԭ��)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrReportDelay_�ӳ�ԭ��) = "", ufgData.ErrCellColor, ufgData.BackColor)
           
    
    
    '���δ¼����ȡҽʦ������ʾ����ɫ
    iCol = ufgData.GetColIndex(gstrReportDelay_�ӳ�����)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrReportDelay_�ӳ�����)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
           
End Sub



Private Sub ShowReasonWindow(ByVal Row As Long, ByVal Col As Long)
    Dim strReason As String
    
    Dim frmReason As New frmPatholReportDelay_Select
    On Error GoTo errFree
    Call frmReason.ShowReasonWindow(ufgData.Text(Row, gstrReportDelay_�ӳ�ԭ��), Me)
    
    strReason = ""
    
    If frmReason.IsOk Then
        With frmReason
            If .chkJF.value <> 0 Then strReason = strReason & IIf(strReason <> "", "����ɷ�", "��ɷ�")
            If .chkTG.value <> 0 Then strReason = strReason & IIf(strReason <> "", "�����Ѹ�", "���Ѹ�")
            If .chkBQC.value <> 0 Then strReason = strReason & IIf(strReason <> "", "���貹ȡ��", "�貹ȡ��")
            If .chkSQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "��������", "������")
            If .chkCQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "��������", "������")
            If .chkLQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "��������", "������")
            If .chkMYZH.value <> 0 Then strReason = strReason & IIf(strReason <> "", "���������黯", "�������黯")
            If .chkFZBL.value <> 0 Then strReason = strReason & IIf(strReason <> "", "������Ӳ���", "����Ӳ���")
            If .chkTSRS.value <> 0 Then strReason = strReason & IIf(strReason <> "", "��������Ⱦɫ", "������Ⱦɫ")
            
            If Trim(.txtOther.Text) <> "" Then strReason = strReason & IIf(strReason <> "", "��", "") & .txtOther.Text
        End With
        
        ufgData.Text(Row, gstrReportDelay_�ӳ�ԭ��) = strReason
    End If
errFree:
    Call Unload(frmReason)
    Set frmReason = Nothing
End Sub



Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowReasonWindow(Row, Col)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub





Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle
    Call LoadReportDelayData(mrecStudyInf.lngPatholAdviceId)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = ufgData.GetColIndex(gstrReportDelay_�ӳ�����) Then
        If Val(ufgData.Text(Row, gstrReportDelay_�ӳ�����)) <= 0 Then ufgData.Text(Row, gstrReportDelay_�ӳ�����) = "1"
        Exit Sub
    End If
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'���浥�Ѵ�ӡ
On Error GoTo errHandle
    Dim strSql As String
    
    strSql = "Zl_�������ӳ�_��ӡ(" & ufgData.KeyValue(ufgData.SelectionRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '�޸Ľ����б�Ĵ�ӡ״̬
    ufgData.Text(ufgData.SelectionRow, gstrReportDelay_��ǰ״̬) = "�Ѵ�ӡ"
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

