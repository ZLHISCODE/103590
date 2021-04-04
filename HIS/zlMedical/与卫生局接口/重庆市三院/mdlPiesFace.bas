Attribute VB_Name = "mdlPiesFace"
Option Explicit

Public glngTXTProc As Long
Public gstrSysName As String
Public gstrDBUser As String
Public gstrSQL As String
Public gstrPrive As String

Public gcnAccess As New Connection

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ���� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type

Public UserInfo As TYPE_USER_INFO

Public Sub OpenRecordSet(rsTemp As ADODB.Recordset, Optional ByVal strFormCaption As String)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close

    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly

End Sub

Public Sub OpenAccessRecordSet(rsTemp As ADODB.Recordset, Optional ByVal strFormCaption As String)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close

    rsTemp.Open gstrSQL, gcnAccess, adOpenStatic, adLockReadOnly

End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("ǰ��ɫ").Name = "ǰ��ɫ")
    blnBkColor = (rsData("����ɫ").Name = "����ɫ")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
        lngRow = 0
    Else
        
        If Val(objMsf.RowData(objMsf.Rows - 1)) <= 0 Then
            lngRow = objMsf.Rows - 2
        Else
            lngRow = objMsf.Rows - 1
        End If
                
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngLoop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If strMask <> "" Then
                    objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
                Else
                    objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(objMsf.TextMatrix(0, lngLoop)))
                End If
            End If
            
        Next
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("ǰ��ɫ").Value)
        If blnBkColor Then objMsf.Cell(flexcpBackColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("����ɫ").Value)
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
    
End Function

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    TxtSelAll objTxt
    objTxt.SetFocus
End Sub



Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function GetCol(ByVal objVsf As Object, ByVal strData As String) As Long
    
    Dim lngLoop As Long
    
    GetCol = -1
    For lngLoop = 0 To objVsf.Cols - 1
        If objVsf.Cell(flexcpText, 0, lngLoop, 0, lngLoop) = strData Then
            GetCol = lngLoop
            Exit Function
        End If
    Next
    
End Function


Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo errHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngLoop = 0 To UBound(varArray)
        varItem = Split(varArray(lngLoop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
'    Exit Function
    
    If objVsf.Rows = 0 Then Exit Function
    lngTop = objVsf.Cell(flexcpTop, objVsf.Rows - 1, 0) + objVsf.RowHeight(objVsf.Rows - 1)
    
    '1.�������е���
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.���¼�����Ҫ������
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function NextNo(intBillID As Integer, Optional ByVal intStep As Integer = 1) As Variant
'���ܣ������ض���������µĺ���,�������£�
'   һ����Ŀ��ţ�
'   1   ����ID         ����
'   2   סԺ��         ����
'   3   �����         ����
'   10  ҽ�����ͺ�     ����,˳��������
'   x   �������ݺ�     �ַ�,���ݱ�Ź���˳��������,���Զ���ȱ
'   �������λȷ��ԭ��:
'       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
    Dim curDate As Date, blnByDate As Boolean
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '����ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            strSQL = "Select Nvl(Max(����ID),0)+1 as ����ID From ������Ϣ Where ����ID>=" & vntNo
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 2 Then 'סԺ��
        '˳���Ż������ڱ��
        strSQL = "Select A.*,Sysdate as ���� From ϵͳ������ A Where A.������=27"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!����ֵ), 1, !����ֵ) = 2)
                curDate = !����
            End If
        End With
        
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(סԺ��),0)+1 as סԺ�� From ������Ϣ Where סԺ��>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(סԺ��),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as סԺ��" & _
                    " From ������Ϣ Where סԺ�� Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And סԺ��>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIf(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 3 Then '�����
        '˳���Ż������ڱ��
        strSQL = "Select A.*,Sysdate as ���� From ϵͳ������ A Where A.������=46"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!����ֵ), 1, !����ֵ) = 2)
                curDate = !����
            End If
        End With
    
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(�����),0)+1 as ����� From ������Ϣ Where �����>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(�����),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as �����" & _
                    " From ������Ϣ Where ����� Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And �����>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "������", IIf(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 10 Then 'ҽ�����ͺ�
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!������), 0, !������)) + 1
            
            On Error Resume Next
            .Update "������", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
        
    ElseIf intBillID = 81 Then      '�����
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!������), 0, !������)) + 1
            
            On Error Resume Next
            .Update "������", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        
        Dim vtnEndNo As Variant
        
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!������), "", !������)
            
            If IIf(IsNull(!��Ź���), 0, !��Ź���) = 1 Then
                '����˳����
                If vntNo < strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000"
                End If
                vtnEndNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + intStep), 4)
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '����˳����
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vtnEndNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + intStep), 7)
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "������", vtnEndNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    'If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetNextPatientID() As Long
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select max(����id) as ����id From ������Ϣ "
    rs.Open strSQL, gcnOracle
    If rs.BOF = False Then
        GetNextPatientID = NVL(rs("����id"), 0)
    End If
    GetNextPatientID = GetNextPatientID + 1
    
End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'����: װ��������ָ�������������������е���������
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Function ShowTxtSelectDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������+�б�ṹ
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo errHand
    
    Call OpenRecord(rs, strSQL, frmParent.Caption, adOpenStatic, adLockBatchOptimistic)
    If rs.BOF Then
        MsgBox "û�п�ѡ������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTxt.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, 3, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rs
        ShowTxtSelectDialog = True
        
    End If
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowTxtFilterDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;��ʾ�ı�����ѡ���б�(ֻ�����ı���ؼ�)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim strInput As String
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand

    If InStr(objTxt.Text, "'") > 0 Then Exit Function
    
    '������ʼ��
    strInput = "'%" & UCase(objTxt.Text) & "%'"
    Call ClientToScreen(objTxt.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        
    'ִ�в�ѯ
    Call OpenRecord(rs, strSQL, frmParent.Caption)
    If rs.BOF Then
        If blnPrompt Then MsgBox "û���ҵ���ƥ��Ľ����", , gstrSysName
        Exit Function                            'û�н����ֱ�ӷ���
    End If
            
    If rs.RecordCount = 1 And blnFilter Then GoTo over                    '��Ϊ��������ң����ֻ��һ������ֱ�ӷ���
    'If frmSelectList.ShowSelect(frmParent, rs, strLvw, lngX, lngY, lngCX, lngCY, strSavePath, strDescrible, , , objTxt.Height) Then GoTo Over
    
    If frmSelectDialog.ShowSelect(frmParent, 2, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, False) Then GoTo over
    
    Exit Function
    
over:
    
    Set rsResult = rs
    
    ShowTxtFilterDialog = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
        
    Set OpenRecord = rsTmp
End Function

Public Sub ExecuteProc(ByVal strSQL As String, ByVal strCaption As String)
'���ܣ�ִ��SQL���
    
    If UCase(Left(strSQL, 3)) = "ZL_" Then
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        gcnOracle.Execute strSQL
    End If
    
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '����:�ر������Ӵ���
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    CloseChildWindows = True
    
    For Each frmThis In Forms
        If frmThis.Caption <> frmMain.Caption And frmThis.Caption <> FrmSon.Caption Then Unload frmThis
    Next
    
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '-----------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��
    '����:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    End Select
    
End Function

Public Function LoadGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("ǰ��ɫ").Name = "ǰ��ɫ")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '��һ��
        Next
        
pointNext:
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("ǰ��ɫ").Value)
        
        rsData.MoveNext
    Loop
    
    LoadGrid = True
    Exit Function
    
errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Public Function SQLInit(ByRef rs As ADODB.Recordset) As Boolean
    
    Set rs = New ADODB.Recordset
    
    With rs
        .Fields.Append "SQL", adVarChar, 4000
        .Open
    End With
    
    SQLInit = True
    
End Function

Public Function SQLAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String) As Boolean
    
    With rs
        .AddNew
        .Fields("SQL").Value = strSQL
    End With
    
    SQLAdd = True
    
End Function

Public Sub ResetVsf(objVsf As Object)
    '
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

Public Function ConnectAccess(ByVal strFile As String) As Boolean
    
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
    Set gcnAccess = New ADODB.Connection
    gcnAccess.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile & ";Persist Security Info=False"
    
    If gcnAccess.State <> adStateOpen Then
        ShowSimpleMsg "����'" & strFile & "'ʧ�ܣ�"
        Exit Function
    End If
    
    ConnectAccess = True
    
End Function

Public Function AcceptPackage(frmMain As Object, ByVal strFile As String, Optional ByVal strTitle As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:���������
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Dim rs As New ADODB.Recordset
    Dim rsTask As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim bytNew As Byte
    Dim intCount As Long
    Dim intCount1 As Long
    Dim intCount2 As Long
    Dim lngTotal As Long
    
    Dim lngKey As Long
    Dim lng����� As Long
    Dim lng����id As Long
            
    If ConnectAccess(strFile) = False Then Exit Function
                
    '1.����
    gstrSQL = "Select taskcode,taskname,taskyear,builddate From htask"
    Call OpenAccessRecordSet(rsTask, strTitle)
    If rsTask.BOF Then Exit Function
         
    '����������Ƿ��ѽ���
    gstrSQL = "Select b.ID,b.���״̬ From ���ǼǼ�¼_�ɱ� a,���ǼǼ�¼ b Where a.�Ǽ�id=b.ID AND a.�������=[1]"
    Set rs = OpenSQLRecord(gstrSQL, strTitle, rsTask("taskcode").Value)
    If rs.BOF = False Then
        
        If NVL(rs("���״̬").Value) >= 4 Then
            ShowSimpleMsg "��������Ѿ���ʼ��죬�������½��ܣ�"
        Else
            If MsgBox("�Ƿ����½��ܵ�ǰ�������", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                GoTo over
            Else
                'ɾ��ԭ����
                lngKey = rs("ID").Value
            End If
        End If
        
    End If
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    frmWait.OpenWait frmMain, "���������"
    
    If lngKey > 0 Then
        frmWait.WaitInfo = "����ɾ��ԭ�����ܵ������..."
        gstrSQL = "ZL_���ǼǼ�¼_DELETE(" & lngKey & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
                
    frmWait.WaitInfo = "���ڽ��������..."
    
    lngKey = GetNextId("���ǼǼ�¼")
    
    
    gstrSQL = "ZL_���ǼǼ�¼_INSERT(" & lngKey & ",'" & _
                                        NextNo(78) & "'," & _
                                        "1," & _
                                        "1," & _
                                        "NULL," & _
                                        "NULL," & _
                                        "NULL," & _
                                        "NULL," & _
                                        Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��Լ��λ", 0)) & "," & _
                                        "1," & _
                                        "Sysdate+1," & _
                                        Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", 0)) & "," & _
                                        "NULL," & _
                                        "Sysdate," & _
                                        "NULL," & _
                                        "1," & _
                                        "1," & _
                                        "NULL)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "Insert Into ���ǼǼ�¼_�ɱ�(�Ǽ�id,�������,�������,����״̬) Values ("
    gstrSQL = gstrSQL & lngKey
    gstrSQL = gstrSQL & ",'" & rsTask("taskcode").Value & "'"
    gstrSQL = gstrSQL & ",'" & rsTask("taskname").Value & "'"
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"
    gcnOracle.Execute gstrSQL
    
    '2.����ײͣ�
    gstrSQL = "Select taskcode,asmcode,asmseq,asmname,asmsex,asmdesc From htaskasm"
    
    Call OpenAccessRecordSet(rs, strTitle)
    If rs.BOF Then GoTo over
    Do While Not rs.EOF
        
        gstrSQL = "ZL_������_INSERT(" & lngKey & ",'" & rs("asmname").Value & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        gstrSQL = "Insert Into ������_�ɱ�(�Ǽ�id,�������,�ײͱ���,�ײ����,�ײ�����) Values ("
        gstrSQL = gstrSQL & lngKey & ","
        gstrSQL = gstrSQL & "'" & rs("asmname").Value & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmcode").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmseq").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmname").Value) & "')"
        gcnOracle.Execute gstrSQL
        
        rs.MoveNext
    Loop
    
    '3.��Ա
    
    frmWait.WaitInfo = "���ڽ��������Ա..."
    frmWait.ShowProgress = True
    
    gstrSQL = "Select a.asmname,b.* From htaskasm a,htaskmemb b Where b.asmcode=a.asmcode and b.asmseq=a.asmseq"
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    Call OpenAccessRecordSet(rs, strTitle)
    
    If rs.BOF Then GoTo over
    
    intCount1 = 0
    lngTotal = rs.RecordCount
    Do While Not rs.EOF
        
        lng����id = 0
        lng����� = 0
        intCount1 = intCount1 + 1
        
        frmWait.WaitProgress = Format(100 * intCount1 / lngTotal, "0.00")
        
        gstrSQL = "Select * From ������Ϣ Where ������='" & rs("membcode").Value & "'"
        Call OpenRecord(rsTmp, gstrSQL, "���������")
        If rsTmp.BOF = False Then
            lng����� = NVL(rsTmp("�����").Value, 0)
            lng����id = NVL(rsTmp("����id").Value, 0)
        End If
        
        bytNew = 0
        If lng����id = 0 Then
            lng����id = GetNextPatientID + intCount
            intCount = intCount + 1
            bytNew = 1
        End If
        
        If lng����� = 0 Then
            intCount2 = intCount2 + 1
            lng����� = NextNo(3) + intCount2
        End If
        
        If rsTmp.BOF = False Then
            gstrSQL = "ZL_�����Ա����_INSERT(" & lngKey & "," & _
                                                lng����id & "," & _
                                                "'" & rs("asmname").Value & "','" & _
                                                rs("a0101").Value & "','" & _
                                                NVL(rsTmp("���֤��").Value) & "','" & _
                                                rs("a0107").Value & "'," & _
                                                IIf(rsTmp("��������").Value = "", "NULL", "TO_DATE('" & rsTmp("��������").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                NVL(rsTmp("����״��").Value) & "','" & _
                                                NVL(rsTmp("����").Value) & "','" & _
                                                NVL(rsTmp("����").Value) & "','" & _
                                                rs("a0405").Value & "','" & _
                                                NVL(rsTmp("ְҵ").Value) & "','" & _
                                                NVL(rsTmp("��ϵ������").Value) & "','" & _
                                                NVL(rsTmp("��ϵ�˵绰").Value) & "','" & _
                                                "','" & _
                                                NVL(rsTmp("��ϵ�˵�ַ").Value) & "','" & _
                                                rs("b0105").Value & "','" & _
                                                rs("age").Value & "'," & _
                                                lng����� & ",'" & _
                                                NVL(rsTmp("IC����").Value) & "','" & _
                                                rs("membcode").Value & "',''," & _
                                                "1," & _
                                                IIf(intCount1 = rs.RecordCount, "1", "0") & ",0," & bytNew & _
                                                ")"
        Else
            gstrSQL = "ZL_�����Ա����_INSERT(" & lngKey & "," & _
                                                lng����id & "," & _
                                                "'" & rs("asmname").Value & "','" & _
                                                rs("a0101").Value & "'," & _
                                                "NULL,'" & _
                                                rs("a0107").Value & "'," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL,'" & _
                                                rs("a0405").Value & "'," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL,'" & _
                                                rs("b0105").Value & "','" & _
                                                rs("age").Value & "'," & _
                                                lng����� & "," & _
                                                "NULL,'" & _
                                                rs("membcode").Value & "',''," & _
                                                "1," & _
                                                IIf(intCount1 = rs.RecordCount, "1", "0") & ",0," & bytNew & _
                                                ")"
        End If
        
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        gstrSQL = "Insert Into �����Ա����_�ɱ�(�Ǽ�id,����id,�������,��Ա���,��λ����,��λ����,��ְ���,��ְ����) Values ("
        gstrSQL = gstrSQL & lngKey & ","
        gstrSQL = gstrSQL & lng����id & ","
        gstrSQL = gstrSQL & "'" & rsTask("taskcode").Value & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("taskseq").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("b0110").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("b0105").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("a6405").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("a0704").Value) & "')"
        gcnOracle.Execute gstrSQL
        
        rs.MoveNext
    Loop
    
    '4.��Ŀ
    frmWait.WaitInfo = "���ڽ��������Ŀ..."
    
    gstrSQL = "Select Distinct a.asmname,b.unioncode From htaskasm a,htaskasmunion b " & _
                "Where a.taskcode=b.taskcode and a.asmcode=b.asmcode and a.asmseq=b.asmseq and a.taskcode='" & rsTask("taskcode").Value & "'"
                
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open gstrSQL, gcnAccess, adOpenStatic
    If rs.BOF Then GoTo over
    
    Dim lngItemKey As Long
    
    intCount1 = 0
    lngTotal = rs.RecordCount
    Do While Not rs.EOF
        
        intCount1 = intCount1 + 1
        
        frmWait.WaitProgress = Format(100 * intCount1 / lngTotal, "0.00")
        
        gstrSQL = "Select b.ID,b.���,0 As �����۸�,0 As ���۸�,0 As ִ�п���id,0 As �ɼ���ʽid,0 As �ɼ�����id,'' As ����걾,'' As �۸��嵥 " & _
                    "From ������ĿĿ¼_�ɱ� a,������ĿĿ¼ b " & _
                    "Where a.������Ŀid=b.ID and a.�ɱ�����='" & rs("unioncode").Value & "'"
        
        Set rsItem = New ADODB.Recordset
        rsItem.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
        
        If rsItem.BOF = False Then
        
            Call FinishFillItem(rsItem, Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", 0)), "��������")
            
            lngItemKey = GetNextId("�����Ŀ�嵥")
            
            gstrSQL = "ZL_�����Ŀ�嵥_INSERT(" & lngKey & "," & _
                                            "'" & rs("asmname").Value & "'," & _
                                            rsItem("ID").Value & "," & _
                                            "NULL," & _
                                            Val(rsItem("�����۸�").Value) & "," & _
                                            Val(rsItem("���۸�").Value) & "," & _
                                            rsItem("ִ�п���id").Value & "," & _
                                            IIf(rsItem("�ɼ���ʽid") = "", "NULL", rsItem("�ɼ���ʽid")) & "," & _
                                            IIf(rsItem("�ɼ�����id") = "", "NULL", rsItem("�ɼ�����id")) & ",'" & _
                                            rsItem("����걾").Value & "'," & _
                                            "NULL," & _
                                            "NULL,NULL, 1,'" & _
                                            NVL(rsItem("�۸��嵥").Value, "") & "')"
            
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            
'            gstrSQL = "Insert Into �����Ŀ�嵥_�ɱ�(�嵥id,�������,��ϱ���,��Ŀ����,��Ŀ��֧,��Ŀ����,��Ͽ���) Values ("
'            gstrSQL = gstrSQL & lngItemKey & ","
'            gstrSQL = gstrSQL & "'" & rsTask("taskcode").Value & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("a6405").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("a0704").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmcode").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmseq").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmname").Value) & "')"
'
'            gcnOracle.Execute gstrSQL
            
        End If
        
        rs.MoveNext
    Loop
    
    gstrSQL = "ZL_���ǼǼ�¼_STATE(" & lngKey & ",2)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    frmWait.CloseWait
    
    gcnOracle.CommitTrans
    
    AcceptPackage = True
    
    Exit Function
    
over:
    frmWait.CloseWait
    gcnOracle.RollbackTrans
    
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    
    frmWait.CloseWait
    
    gcnOracle.RollbackTrans
    
    ShowSimpleMsg strError
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
'    Resume
End Function

Private Function FinishFillItem(ByRef rsItem As ADODB.Recordset, ByVal mlngDept As Long, Optional ByVal strTitle As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strKeys As String
    Dim sglSum As Single
    Dim lngExecDept As Long
    Dim strTmp As String
    Dim lngLoop As Long
    Dim strCombList As String
    Dim lngKey As Long
        
    On Error GoTo errHand
    
    lngKey = rsItem("ID").Value
    
    'ִ�п���id
    gstrSQL = GetPublicSQL(SQL.����ִ�п���)
    If gstrSQL <> "" Then
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey, mlngDept, UserInfo.����ID, "%%")
        If rs.BOF = False Then
            rsItem("ִ�п���id").Value = NVL(rs("ID").Value)
        End If
    End If
        
    If rsItem("���").Value = "C" Then
        '�ɼ���ʽid
        gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            rsItem("�ɼ���ʽid").Value = NVL(rs("ID").Value)
        Else
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A WHERE A.���='E' AND A.��������='6'"
            Set rs = OpenSQLRecord(gstrSQL, strTitle)
            If rs.BOF = False Then
                rsItem("�ɼ���ʽid").Value = NVL(rs("ID").Value)
            End If
        End If
            
        '�ɼ�����id
        gstrSQL = GetPublicSQL(SQL.����ִ�п���)
        Set rs = OpenSQLRecord(gstrSQL, strTitle, Val(rsItem("�ɼ���ʽid").Value), mlngDept, UserInfo.����ID, "%%")
        If rs.BOF = False Then
            rsItem("�ɼ�����id").Value = NVL(rs("ID").Value)
        End If
        
        
        '����걾
        gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            '�������Ŀ
            
            gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                    "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                        "AND B.������Ŀid=A.��Ŀid and rownum<2"
                        
            gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                         "FROM ���鱨����Ŀ A," & _
                              "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                              "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                        "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                              "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)  and rownum<2 "
                                      
        Else
            gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                    "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid  and rownum<2"
        End If
        
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            rsItem("����걾").Value = rs("����").Value
        Else
            
            'û�ж�Ӧʱ����ȡ���б걾����
            gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A where rownum<2"
            Set rs = OpenSQLRecord(gstrSQL, strTitle)
            If rs.BOF = False Then
                rsItem("����걾").Value = rs("����").Value
            End If
            
        End If
    End If
    
    '�۸�
        
    strKeys = rsItem("ID").Value & "'" & rsItem("�ɼ���ʽid").Value & "'0"
    
    gstrSQL = GetPublicSQL(SQL.�����Ŀ�۱�, strKeys)
    Set rs = OpenSQLRecord(gstrSQL, strTitle)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            
            sglSum = sglSum + NVL(rs("�շ�����"), 0) * NVL(rs("�ּ�"), 0)
                        
            '�շ�ִ�п���
            If InStr("4,5,6,7", rs("���").Value) > 0 Then
                gstrSQL = GetPublicSQL(SQL.ҩƷִ�п���)
                Set rsTmp = OpenSQLRecord(gstrSQL, strTitle, rs("���").Value)
            Else
                gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���)
                Set rsTmp = OpenSQLRecord(gstrSQL, strTitle, lngKey, mlngDept, UserInfo.����ID, "%%")
            End If
            
            If rsTmp.BOF = False Then
                lngExecDept = NVL(rsTmp("ID").Value)
            Else
                lngExecDept = rsItem("ִ�п���id").Value
            End If
            
            strTmp = strTmp & ";" & NVL(rs("ID")) & ":" & NVL(rs("�շ�����")) & ":" & NVL(rs("�ּ�")) & ":" & NVL(rs("�ּ�")) & ":" & lngExecDept & ":" & NVL(rs("�Ƽ�����"))
            
            
            rs.MoveNext
        Loop
    End If
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    rsItem("�۸��嵥").Value = strTmp
    rsItem("�����۸�").Value = sglSum
    rsItem("���۸�").Value = sglSum
    
    FinishFillItem = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
'    Resume
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = ActualLen(varValue)
            If intMax = 0 Or intMax < 10 Then intMax = 10
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = ActualLen(varValue(lngLeft))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Set OpenSQLRecord = cmdData.Execute
    
End Function

Public Function ImportData(frmMain As Object, _
                            ByVal strFile As String, _
                            ByVal lng���Ʒ���id As Long, _
                            ByVal str���Ʒ������ As String, _
                            ByVal lng���η���id As Long, _
                            ByVal str���η������ As String, _
                            Optional ByVal lng����ִ�п��� As Long) As Boolean
    
    '------------------------------------------------------------------------------------------------------------------
    '����:�����������ṩ�������Ŀ�������ж����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strSvrCode As String
    Dim lng������Ŀid As Long
    Dim lng������Ŀid As Long
    Dim lng�����Ŀid As Long
    Dim lngNo As Long
    Dim lngNo2 As Long
    Dim lngTotal As Long
    Dim lngElementID As Long
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim lngCount As Long
    Dim byt���� As Byte
    
    ReDim Preserve strSQL(1 To 1)
    
    On Error GoTo errHand
    
    If ConnectAccess(strFile) = False Then Exit Function
    
    gcnOracle.BeginTrans
    
    frmWait.OpenWait frmMain, "ת�������Ŀ"
    
    frmWait.WaitInfo = "����ɾ��ԭ������..."
    
    '�����Ӧ
    gstrSQL = "Delete From ���鱨����Ŀ Where ������Ŀid In (Select ID From ������ĿĿ¼ Where ����id=" & lng���Ʒ���id & ")"
    gcnOracle.Execute gstrSQL
    
    '����Ӧ
    gstrSQL = "Delete From ����Ԫ��Ŀ¼ Where ����=-1 AND ����='00000'"
    gcnOracle.Execute gstrSQL
    
    gstrSQL = "Delete From ����������Ŀ_�ɱ�"
    gcnOracle.Execute gstrSQL
    
    '����
    gstrSQL = "Delete From ����������Ŀ Where ����id=" & lng���η���id
    gcnOracle.Execute gstrSQL
            
    '�����
    gstrSQL = "Delete From ����������Ŀ Where ����id IS NULL AND ���� In (Select ���� From ������ĿĿ¼ Where ���='C' AND ����id=" & lng���Ʒ���id & ")"
    gcnOracle.Execute gstrSQL
    
    
    gstrSQL = "Delete From ������ĿĿ¼_�ɱ�"
    gcnOracle.Execute gstrSQL
    
    '����
    gstrSQL = "Delete From ������ĿĿ¼ Where ����id=" & lng���Ʒ���id
    gcnOracle.Execute gstrSQL
    
        
    frmWait.WaitInfo = "���ڴ�����..."
            
    gstrSQL = "Select * From �����Ŀö��ֵ Order By �����Ŀ��,��"
    If rsData.State = adStateOpen Then rsData.Close
    rsData.CursorLocation = adUseClient
    rsData.Open gstrSQL, gcnAccess, adOpenStatic
    
    gstrSQL = "Select * From �����Ŀ_���� Where ��ϱ��� Not In ('R80') Order By ��ϱ���,��Ŀ����"
    
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open gstrSQL, gcnAccess, adOpenStatic
    If rs.BOF = False Then
        
        frmWait.WaitInfo = "����ת������..."
        frmWait.ShowProgress = True
        
        lngTotal = rs.RecordCount
        lngLoop = 0
        
        lngElementID = GetNextId("����Ԫ��Ŀ¼")
        gstrSQL = "ZL_����Ԫ��_INSERT(-1," & lngElementID & ",'00000','������Ӧ','','����,9',1,Null,'00001')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        Do While Not rs.EOF
            
            lngLoop = lngLoop + 1
            frmWait.WaitProgress = Format(100 * lngLoop / lngTotal, "0.00")
            
            If strSvrCode <> rs("��ϱ���").Value Then
                
                lngCount = lngCount + 1
                
                strSvrCode = rs("��ϱ���").Value
                
                lng�����Ŀid = GetNextId("������ĿĿ¼")
                                
                gstrSQL = "zl_������Ŀ_Insert('"
                gstrSQL = gstrSQL & IIf(UCase(Left(rs("��ϱ���").Value, 1)) = "R", "C", "D") & "',"
                gstrSQL = gstrSQL & lng���Ʒ���id & ","
                gstrSQL = gstrSQL & lng�����Ŀid & ",'"
                gstrSQL = gstrSQL & (str���Ʒ������ & Format(lngCount, "0000")) & "','"
                gstrSQL = gstrSQL & rs("�������").Value & "','"
                gstrSQL = gstrSQL & zlGetSymbol(rs("�������").Value, 0) & "','"                    '����ƴ��_IN ������Ŀ����.����%TYPE := NULL,
                gstrSQL = gstrSQL & zlGetSymbol(rs("�������").Value, 1) & "',"                    '�������_IN ������Ŀ����.����%TYPE := NULL,
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,'"
                gstrSQL = gstrSQL & "����',"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "3,'"
                gstrSQL = gstrSQL & rs("��λ").Value & "',"
                gstrSQL = gstrSQL & "0,"
                gstrSQL = gstrSQL & "0,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "'',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "4,"
                If UCase(Left(rs("��ϱ���").Value, 1)) = "R" Then
                    gstrSQL = gstrSQL & IIf(lng����ִ�п��� = 0, "NULL", lng����ִ�п���) & ","         '����ִ�п���
                Else
                    gstrSQL = gstrSQL & "NULL,"             '����ִ�п���
                End If
                gstrSQL = gstrSQL & "NULL,"             'סԺִ�п���
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "0)"
                
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                gstrSQL = "Insert Into ������ĿĿ¼_�ɱ�(������Ŀid,�ɱ�����,�ɱ�����,��Ͽ���) Values ("
                gstrSQL = gstrSQL & lng�����Ŀid & ","
                gstrSQL = gstrSQL & "'" & rs("��ϱ���").Value & "',"
                gstrSQL = gstrSQL & "'" & rs("�������").Value & "',"
                gstrSQL = gstrSQL & "'" & rs("���ұ���").Value & "')"
                
                gcnOracle.Execute gstrSQL
                
                lngNo = 0
                lngNo2 = 0
            End If
                        
            If NVL(rs("����").Value) <> "" Or NVL(rs("����").Value) <> "" Then
                byt���� = 0
                '����
            Else
                '�ı�
                byt���� = 1
            End If
            
            If UCase(Left(rs("��ϱ���").Value, 1)) = "R" Then
                
                '������Ŀ
                lngCount = lngCount + 1
                
                lng������Ŀid = GetNextId("������ĿĿ¼")
                
                gstrSQL = "Select a.������Ŀid From ����������Ŀ_�ɱ� a Where a.�ɱ�����=[1]"
                Set rsTmp = OpenSQLRecord(gstrSQL, "ת������", rs("��Ŀ����").Value)
                If rsTmp.BOF Then
                    gstrSQL = "zl_������Ŀ_Insert('"
                    gstrSQL = gstrSQL & "C',"                                                       '���
                    gstrSQL = gstrSQL & lng���Ʒ���id & ","                                         '����ID
                    gstrSQL = gstrSQL & lng������Ŀid & ",'"
                    gstrSQL = gstrSQL & (str���Ʒ������ & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("��Ŀ����").Value & "','"                                '����
                    gstrSQL = gstrSQL & zlGetSymbol(rs("��Ŀ����").Value, 0) & "','"
                    gstrSQL = gstrSQL & zlGetSymbol(rs("��Ŀ����").Value, 1) & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & "����',"
                    gstrSQL = gstrSQL & "1,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "3,'"
                    gstrSQL = gstrSQL & rs("��λ").Value & "',"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "1,"
                    gstrSQL = gstrSQL & "0,"                '�����Ŀ
                    gstrSQL = gstrSQL & "'',"               '�걾��λ
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "4,"                'ִ�п���
                    gstrSQL = gstrSQL & IIf(lng����ִ�п��� = 0, "NULL", lng����ִ�п���) & ","         '����ִ�п���
                    gstrSQL = gstrSQL & "NULL,"             'סԺִ�п���
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "0)"
                    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    lng������Ŀid = GetNextId("����������Ŀ")
    
                    gstrSQL = "ZL_������Ŀ_INSERT("
                    gstrSQL = gstrSQL & lng������Ŀid & ","
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & (str���Ʒ������ & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("��Ŀ����").Value & "',"                             '������
                    gstrSQL = gstrSQL & "NULL,"                                                 'Ӣ����
                    gstrSQL = gstrSQL & byt���� & ","                                                    '����
                    gstrSQL = gstrSQL & "50,"
                    gstrSQL = gstrSQL & "0,'"
                    gstrSQL = gstrSQL & rs("��λ").Value & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL)"
    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    gstrSQL = "ZL_������Ŀ_UPDATE("
                    gstrSQL = gstrSQL & lng������Ŀid & ","
                    gstrSQL = gstrSQL & "NULL,"                     '��д
                    gstrSQL = gstrSQL & "NULL,"                     '�������
                    gstrSQL = gstrSQL & "1,"                        '��Ŀ���
                    gstrSQL = gstrSQL & IIf(byt���� = 0, 1, 2) & ",'"                   '�������
                    gstrSQL = gstrSQL & rs("��λ").Value & "',"     '��λ
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL)"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    gstrSQL = "ZL_���鱨����Ŀ_UPDATE(" & lng������Ŀid & ",'^" & lng������Ŀid & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                                        
                    gstrSQL = "Insert Into ����������Ŀ_�ɱ�(������Ŀid,�ɱ�����,�ɱ�����,��Ŀ��֧,��Ŀ����) VALUES ("
                    gstrSQL = gstrSQL & lng������Ŀid & ","
                    gstrSQL = gstrSQL & "'" & rs("��Ŀ����").Value & "',"
                    gstrSQL = gstrSQL & "'" & rs("��Ŀ����").Value & "',"
                    gstrSQL = gstrSQL & "'',"
                    gstrSQL = gstrSQL & "'')"
                    gcnOracle.Execute gstrSQL
                    
                    gstrSQL = "ZL_������Ŀȡֵ_DELETE(" & lng������Ŀid & ")"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    '���ҿ�ѡ����,�п�ѡ����
                    rsData.Filter = ""
                    rsData.Filter = "�����Ŀ��='" & NVL(rs("��Ŀ����").Value) & "'"
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                        Do While Not rsData.EOF
                            
                            gstrSQL = "ZL_������Ŀȡֵ_INSERT(" & lng������Ŀid & ",'" & NVL(rsData("��").Value) & "','" & NVL(rsData("ö��ֵ").Value) & "',0)"
                            gcnOracle.Execute gstrSQL, , adCmdStoredProc
                            
                            rsData.MoveNext
                        Loop
                    End If
                    
                    '�����������
                    If byt���� = 0 Then
                        
                        gstrSQL = "ZL_������Ŀ�ο�_DELETE(" & lng������Ŀid & ")"
                        gcnOracle.Execute gstrSQL, , adCmdStoredProc
                        
                        gstrSQL = "ZL_������Ŀ�ο�_INSERT(" & lng������Ŀid & ",'',0,NULL,NULL,NULL," & Val(NVL(rs("����").Value)) & "," & Val(NVL(rs("����").Value)) & ",'')"
                        gcnOracle.Execute gstrSQL, , adCmdStoredProc
                        
                    End If
                Else
                    lng������Ŀid = rsTmp("������Ŀid").Value
                End If
                
                lngNo = lngNo + 1
                gstrSQL = "insert into ���鱨����Ŀ(������ĿID,����걾,������ĿID,�������) values (" & lng�����Ŀid & ",NULL," & lng������Ŀid & "," & lngNo & ")"
                gcnOracle.Execute gstrSQL
                
            Else
                
                '��д������Ŀ
                
                lngCount = lngCount + 1
                
                gstrSQL = "Select a.������Ŀid From ����������Ŀ_�ɱ� a Where a.�ɱ�����=[1]"
                Set rsTmp = OpenSQLRecord(gstrSQL, "ת������", rs("��Ŀ����").Value)
                If rsTmp.BOF Then
                    
                    strTmp = ""
                    
                    '���ҿ�ѡ����,�п�ѡ���ݣ��Ͷ���Ϊ����ѡ���;����Ϊ�ı������
                    If byt���� = 1 Then
                        rsData.Filter = ""
                        rsData.Filter = "�����Ŀ��='" & NVL(rs("��Ŀ����").Value) & "'"
                        If rsData.RecordCount > 0 Then
                            rsData.MoveFirst
                            Do While Not rsData.EOF
                                strTmp = strTmp & ";" & rsData("ö��ֵ").Value
                                rsData.MoveNext
                            Loop
                            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                        End If
                    Else
                        strTmp = Val(NVL(rs("����").Value)) & ";" & Val(NVL(rs("����").Value))
                    End If
                    
                    lng������Ŀid = GetNextId("����������Ŀ")
                    
                    gstrSQL = "ZL_������Ŀ_INSERT("
                    gstrSQL = gstrSQL & lng������Ŀid & ","
                    gstrSQL = gstrSQL & lng���η���id & ",'"
                    gstrSQL = gstrSQL & (str���η������ & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("��Ŀ����").Value & "',"                             '������
                    gstrSQL = gstrSQL & "NULL,"                                                 'Ӣ����
                    gstrSQL = gstrSQL & byt���� & ","                                           '����
                    gstrSQL = gstrSQL & "50,"
                    gstrSQL = gstrSQL & "0,'"
                    gstrSQL = gstrSQL & rs("��λ").Value & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & IIf(strTmp <> "" And byt���� = 1, "2", "0") & ","                       '��ʾ��
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & strTmp & "',"                                           '��ֵ��
                    gstrSQL = gstrSQL & "NULL,"                                                 '��ʼֵ
                    gstrSQL = gstrSQL & "NULL,"                                                 '���ֱ���
                    gstrSQL = gstrSQL & "NULL)"                                                 '��ֵ����
    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
                    gstrSQL = "Insert Into ����������Ŀ_�ɱ�(������Ŀid,�ɱ�����,�ɱ�����,��Ŀ��֧,��Ŀ����) VALUES ("
                    gstrSQL = gstrSQL & lng������Ŀid & ","
                    gstrSQL = gstrSQL & "'" & rs("��Ŀ����").Value & "',"
                    gstrSQL = gstrSQL & "'" & rs("��Ŀ����").Value & "',"
                    gstrSQL = gstrSQL & "'',"
                    gstrSQL = gstrSQL & "'')"
                    gcnOracle.Execute gstrSQL
                Else
                    lng������Ŀid = rsTmp("������Ŀid").Value
                End If
                
                lngNo2 = lngNo2 + 1
  
                gstrSQL = "ZL_������_SAVE("
                gstrSQL = gstrSQL & lngElementID & ","
                gstrSQL = gstrSQL & lngNo2 & ","
                gstrSQL = gstrSQL & "'2',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & lng�����Ŀid & ","
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & lng������Ŀid & ","
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "'" & NVL(rs("��λ").Value) & "',"                  '��λ
                gstrSQL = gstrSQL & "NULL)"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
            End If
            
            rs.MoveNext
        Loop
            
    End If
    
    frmWait.ShowProgress = False
    frmWait.WaitInfo = "���ڱ�������..."
    
        
    gcnOracle.CommitTrans
    
    frmWait.CloseWait
    
    ImportData = True
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    frmWait.CloseWait
    gcnOracle.RollbackTrans
    ShowSimpleMsg strError
    
'    Resume
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    
    zlGetSymbol = "-"
End Function




