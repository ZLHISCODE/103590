Attribute VB_Name = "mdlClinicPlanFun"
Option Explicit
'�ҺŰ��Ź����嵥
Public Enum RegistPlanFun
    Pane_FunFace = 1
    Pane_Face = 2
    
    Pane_WorkTime = 11 '����ʱ�����
    Pane_Holiday = 12 '�ڼ��չ���
    Pane_DoctorOffice = 13 '������������
    Pane_SignalSource = 14 '��Դ������
    
    Pane_StopPlan = 21 'ͣ�����
    Pane_FixedPlan = 22 '�̶�����
    Pane_PlanTemplet = 23 '����ģ��
    Pane_MonthPlan = 24 '�°���
    Pane_WeekPlan = 25 '�ܰ���
    Pane_MonthTemplet = 26 '�������ó��ﰲ�ŵ���ģ��
End Enum

Public Sub ZlUpdatePlanMenu(frmParent As Object, cbsMain As Object, Optional ByVal bytFun As Byte, Optional ByVal lng��ԱID As Long)
    '���ò˵�����
    Dim cbrControl As CommandBarControl
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim strMonthMenu As String, strWeekMenu As String
    
    On Error Resume Next
    If cbsMain Is Nothing Then Exit Sub
    If GetNextPlanDate(frmParent, 1, intYear, intMonth, intWeek, lng��ԱID, False) = False Then Exit Sub
    If intMonth = 0 Then
        strMonthMenu = "�����³����"
    Else
        strMonthMenu = "����" & intMonth & "�³����"
    End If
    If bytFun = 1 Then
        Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu & "(&N)"
        Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu
    End If
    
    Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextMonthNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu & "(&N)"
    Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextMonthNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu
    
    If GetNextPlanDate(frmParent, 2, intYear, intMonth, intWeek, lng��ԱID, False) = False Then Exit Sub
    If intWeek = 0 Then
        strWeekMenu = "�����ܳ����"
    Else
        strWeekMenu = "����" & intMonth & "�µ�" & intWeek & "�ܳ����"
    End If
    If bytFun <> 1 Then
        Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu & "(&N)"
        Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu
    End If
    
    Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextWeekNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu & "(&W)"
    Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextWeekNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu
    cbsMain.RecalcLayout
End Sub

Public Function CheckIsHavePlan(ByVal byt�Ű෽ʽ As Byte, ByVal lngUserID As Long, _
    Optional ByVal dt��ʼʱ�� As Date, Optional ByVal dt��ֹʱ�� As Date, _
    Optional ByVal blnDeleteFixedPlan As Boolean) As Boolean
    '��鵱ǰ����Ա�Ƿ��пɲ����ĺ�Դ
    '��Σ�
    '   byt�Ű෽ʽ - 0:�̶��Ű�,1:���Ű�,2:���Ű�,3:ģ��
    '   lngUserID - �û�ID
    '   dt��ʼʱ�䡢dt��ֹʱ�� -  �³�����ʱ�䷶Χ
    '   blnDeleteFixedPlan - �Ƿ�ɾ���޹Һ�ԤԼ�ĳ����¼
    Dim strWhere As String, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lngUserID > 0 Then
        strWhere = " And Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [2])"
    End If
    Select Case byt�Ű෽ʽ
    Case 0 '�̶������
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D" & vbNewLine & _
                " Where a.����ID = b.ID And a.ҽ��ID = c.ID(+) And a.��ĿID = d.ID And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                "       And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null)" & vbNewLine & _
                "       And Nvl(Nvl(b.վ��, [3]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                "       And Not Exists(Select 1 From �ٴ����ﰲ�� P,�ٴ������ Q" & vbNewLine & _
                "                      Where p.����ID = q.ID And p.��ԴID = a.ID And q.�Ű෽ʽ = 0)" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ", gstrNodeNo, lngUserID, gVisitPlan_ModulePara.str��Դά��վ��)
    Case 1, 2
        If byt�Ű෽ʽ = 1 Then
            strWhere = " And �Ű෽ʽ = 1" & vbNewLine & strWhere
        Else
            '1.��ǰ��ԴΪ���Ű����ڳ����ʱ�䷶Χ���޳����¼
            '2.��ǰ��Դ�������Ű����Ϊ�����Ű࣬����ǰ���ڳ�����������������Ű࣬������������ʣ�µĲ��ֽ��������ܽ����Ű�
            strWhere = " And (a.�Ű෽ʽ = 2 And Not Exists (Select 1" & vbNewLine & _
                    "           From �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                    "           Where p.����id = q.Id And p.��Դid = a.Id" & vbNewLine & _
                    "               And Not(p.��ֹʱ��<[5] Or p.��ʼʱ��>Last_Day([3])) And q.�Ű෽ʽ = 1)" & vbNewLine & _
                    "       Or a.�Ű෽ʽ = 1 And Exists (Select 1" & vbNewLine & _
                    "           From �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                    "           Where p.����id = q.Id And p.��Դid = a.Id" & vbNewLine & _
                    "               And Not(p.��ֹʱ��<[5] Or p.��ʼʱ��>Last_Day([3])) And q.�Ű෽ʽ = 2))" & vbNewLine & _
                    strWhere
        End If
        
        '��Դ�ڸó����ʱ�䷶Χ���޳����¼
        strWhere = _
                "  And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "        From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q" & vbNewLine & _
                "        Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id" & _
                "              And o.�������� Between [3] And [4] " & vbNewLine & _
                         IIf(blnDeleteFixedPlan, _
                "              And (q.�Ű෽ʽ In (1, 2) Or q.�Ű෽ʽ = 0 And (Nvl(o.�ѹ���, 0) <> 0 Or Nvl(o.��Լ��, 0) <> 0))", "") & ")" & vbNewLine & _
                strWhere
        
        strSQL = "Select 1" & vbNewLine & _
            " From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D" & vbNewLine & _
            " Where a.����ID = b.ID And a.ҽ��ID = c.ID(+) And a.��ĿID = d.ID And Nvl(a.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
            "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
            "       And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
            "       And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(b.վ��, [6]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                    strWhere & vbNewLine & _
            "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ", gstrNodeNo, lngUserID, dt��ʼʱ��, dt��ֹʱ��, _
            CDate(Format(dt��ʼʱ��, "yyyy-mm-01")), gVisitPlan_ModulePara.str��Դά��վ��)
    Case 3 'ģ��
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D" & vbNewLine & _
                " Where a.����ID = b.ID And a.ҽ��ID = c.ID(+) And a.��ĿID = d.ID And a.�Ű෽ʽ In (1, 2) And Nvl(a.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                "       And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null)" & vbNewLine & _
                "       And Nvl(Nvl(b.վ��, [3]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ", gstrNodeNo, lngUserID, gVisitPlan_ModulePara.str��Դά��վ��)
    End Select
    CheckIsHavePlan = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckExistsFixedToOthers(ByVal byt�Ű෽ʽ As Byte, ByVal lngUserID As Long, _
     ByVal dt��ʼʱ�� As Date, ByRef strDelInfo As String, ByRef strNotDelInfo As String) As Boolean
    '��������Ű෽ʽ�ĺ�Դ�Ƿ����ɹ̶��Ű�ת��������
    '��Σ�
    '   byt�Ű෽ʽ - 1:���Ű�,2:���Ű�
    '   lngUserID - �û�ID
    '   dt��ʼʱ�� -  �³����Ŀ�ʼʱ��
    '����
    '   strDelInfo - ��ɾ�������¼���ڵ�ǰ�Ű෽ʽ�ĺ�Դ,��ʽ������-����(ҽ������) + vbCrLf + ����-����(ҽ������) + vbCrLf + ...
    '   strNotDelInfo - ����ɾ�������¼�������ڵ�ǰ�Ű෽ʽ�ĺ�Դ,��ʽ������-����(ҽ������) + vbCrLf + ����-����(ҽ������) + vbCrLf + ...
    Dim strWhere As String, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strDelInfo = "": strNotDelInfo = ""
    If lngUserID > 0 Then
        strWhere = " And Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = d.����id And ��Աid = [4])"
    End If
    strSQL = "Select Max(Decode(e.Id, Null, 0, 1)) As ����, d.����, Max(f.����) As ����, Max(d.ҽ������) As ҽ������" & vbNewLine & _
            " From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D, ���˹Һż�¼ E, ���ű� F" & vbNewLine & _
            " Where a.����id = b.Id And b.����id = c.Id And b.��Դid = d.Id And a.Id = e.�����¼id(+) And d.����id = f.Id And c.�Ű෽ʽ = 0" & vbNewLine & _
            "       And a.�������� >= [3] And Nvl(d.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
            "       And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(d.�Ű෽ʽ, 0) = [2]" & vbNewLine & _
            "       And Nvl(Nvl(f.վ��,[5]),Nvl([1],'-')) = Nvl([1],'-')" & vbNewLine & _
                    strWhere & _
            " Group By d.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Դ", gstrNodeNo, byt�Ű෽ʽ, dt��ʼʱ��, lngUserID, gVisitPlan_ModulePara.str��Դά��վ��)
    Do While Not rsTemp.EOF
        If Val(Nvl(rsTemp!����)) = 0 Then
            strDelInfo = strDelInfo & vbCrLf & "  " & Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) & IIf(Nvl(rsTemp!ҽ������) = "", "", "(" & Nvl(rsTemp!ҽ������) & ")")
        End If
        If Val(Nvl(rsTemp!����)) = 1 Then
            strNotDelInfo = strNotDelInfo & vbCrLf & "  " & Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) & IIf(Nvl(rsTemp!ҽ������) = "", "", "(" & Nvl(rsTemp!ҽ������) & ")")
        End If
        rsTemp.MoveNext
    Loop
    If strDelInfo <> "" Then strDelInfo = Mid(strDelInfo, 4)
    If strNotDelInfo <> "" Then strNotDelInfo = Mid(strNotDelInfo, 4)
    CheckExistsFixedToOthers = Not (strDelInfo = "" And strNotDelInfo = "")
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlDeletePlan(ByVal lng����ID As Long, Optional ByVal lngUserID As Long) As Boolean
    '���ܣ�ɾ�������
    '��Σ�
    '   lngUserID - �û�ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    
    'Zl_�ٴ������_Delete
    strSQL = "Zl_�ٴ������_Delete("
    '  Id_In       �ٴ������.Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '  ��Աid_In ��Ա��.Id%Type := Null
    strSQL = strSQL & "" & lngUserID & ","
    '  վ��_In   ���ű�.վ��%Type
    strSQL = strSQL & "'" & gstrNodeNo & "')"
    zlDatabase.ExecuteProcedure strSQL, "ɾ�������"
    
    ZlDeletePlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlClearPlan(ByVal lng����ID As Long, ByVal strItem As String, _
    Optional ByVal blnRecord As Boolean) As Boolean
    '���ܣ����ĳһ��İ���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Or strItem = "" Then Exit Function
    
    'Zl_�ٴ������ϰ�ʱ��_Delete(
    strSQL = "Zl_�ٴ������ϰ�ʱ��_Delete("
    '����id_In   �ٴ���������.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '��Ŀ_In     �ٴ���������.������Ŀ%Type,
    strSQL = strSQL & "'" & strItem & "',"
    '�����¼_In Number := 0,
    strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
    '�ϰ�ʱ��_In     �ٴ���������.�ϰ�ʱ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    'ɾ�����ﰲ��_In Number:=0
    strSQL = strSQL & "" & 1 & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "�������"
    ZlClearPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlClearPlanBatch(ByVal lng����ID As Long, Optional ByVal lng��ԴId As Long, _
    Optional ByVal lng��ԱID As Long, Optional ByVal lng����ID As Long, Optional ByVal blnTempPlan As Boolean) As Boolean
    '���ܣ�������к�Դ���ţ�����ɾ��ĳһ����Դ�����а���
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function

    'Zl_�ٴ����ﰲ��_BatchDelete(
    strSQL = "Zl_�ٴ����ﰲ��_BatchDelete("
    '����id_In �ٴ������.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '��Աid_In ��Ա��.Id%Type := 0,
    strSQL = strSQL & "" & lng��ԱID & ","
    'վ��_In   ���ű�.վ��%Type := Null,
    strSQL = strSQL & "" & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '��Դid_In �ٴ����ﰲ��.��Դid%Type := 0
    strSQL = strSQL & "" & lng��ԴId & ","
    '����id_In �ٴ����ﰲ��.Id%Type := 0
    strSQL = strSQL & "" & lng����ID & ","
    '��ʱ����_In �ٴ����ﰲ��.�Ƿ���ʱ����%Type := 0
    strSQL = strSQL & "" & IIf(blnTempPlan, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, "�����������"
    ZlClearPlanBatch = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlPlanApplyTo(ByVal bytType As Byte, ByVal lngԭ����ID As Long, ByVal strԭ��Ŀ As String, _
    ByVal lng�°���ID As Long, ByVal str����Ŀ As String, Optional ByVal blnTemp As Boolean) As Boolean
    '���ܣ�Ӧ������������
    '������
    '   bytType 0-ģ���̶������,1-�����¼
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    If lngԭ����ID = 0 Or lng�°���ID = 0 _
        Or strԭ��Ŀ = "" Or str����Ŀ = "" Then Exit Function
    
    'Zl_�ٴ����ﰲ��_Applyto(
    strSQL = "Zl_�ٴ����ﰲ��_Applyto("
    'Ӧ������_In     Number,--0-ģ���̶������,1-�����¼
    strSQL = strSQL & "" & bytType & ","
    'ԭid_In         �ٴ����ﰲ��.Id%Type,
    strSQL = strSQL & "" & lngԭ����ID & ","
    'ԭ��Ŀ_In       Varchar2,
    strSQL = strSQL & "'" & strԭ��Ŀ & "',"
    '��id_In         �ٴ����ﰲ��.Id%Type,
    strSQL = strSQL & "" & lng�°���ID & ","
    '����Ŀ_In       Varchar2,--Ӧ���ڵ���Ŀ�������"|"�ָ���
    strSQL = strSQL & "'" & str����Ŀ & "',"
    '�Ƿ���ʱ����_In Number:=0
    strSQL = strSQL & "" & IIf(blnTemp, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, "Ӧ������������"
    ZlPlanApplyTo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlBatchSNControl(ByVal lng����ID As Long, ByVal bytStart As Boolean, _
    Optional ByVal lng��ԱID As Long) As Boolean
    '���ܣ�ȫ��������ſ��ƻ���ȫ��ȡ����ſ���
    '������
    '   bytStart True-����,False-ͣ��
    '   blnRecord True-�����¼,False-��������
    Dim strSQL As String

    On Error GoTo errHandler
    If lng����ID = 0 Then Exit Function
    'Zl_�ٴ����ﰲ��_��ſ���(
    strSQL = "Zl_�ٴ����ﰲ��_��ſ���("
    '����id_In   �ٴ������.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '��ſ���_In �ٴ���������.�Ƿ���ſ���%Type,
    strSQL = strSQL & "" & IIf(bytStart, 1, 0) & ","
    'վ��_In     ���ű�.վ��%Type := Null,
    strSQL = strSQL & "" & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '��Աid_In   ��Ա��.Id%Type := 0
    strSQL = strSQL & "" & ZVal(lng��ԱID) & ")"
    zlDatabase.ExecuteProcedure strSQL, "������ſ���"
    
    ZlBatchSNControl = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZlBatchLockPlan(ByVal str��¼IDs As String, ByVal blnUnlock As Boolean) As Boolean
    '���������¼
    '��Σ�
    '   blnUnlock �Ƿ����,True-����,False-����
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If str��¼IDs = "" Then Exit Function
    If blnUnlock Then
        'Zl_�ٴ������¼_Batchlock
        '  -- Ids_In �������������������ö��ŷָ�
        strSQL = "Zl_�ٴ������¼_Batchlock("
        '  Ids_In      Varchar2,
        strSQL = strSQL & "'" & str��¼IDs & "',"
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
        zlDatabase.ExecuteProcedure strSQL, "��������"
    Else
        'Zl_�ٴ������¼_Batchlock
        '  -- Ids_In �������������������ö��ŷָ�
        strSQL = "Zl_�ٴ������¼_Batchlock("
        '  Ids_In      Varchar2,
        strSQL = strSQL & "'" & str��¼IDs & "')"
        '  ȡ������_In Number:=0
        zlDatabase.ExecuteProcedure strSQL, "��������"
    End If
    ZlBatchLockPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistsPlanOnVisitTable(ByVal byt�Ű෽ʽ As Byte, ByVal lng����ID As Long, _
    Optional ByVal lngUserID As Long) As Boolean
    '��鵱ǰ��/�ܳ�������Ƿ������Ч�İ���
    '��Σ�
    '   byt�Ű෽ʽ 1-���Ű�,2-���Ű�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From �ٴ������Դ A, �ٴ����ﰲ�� B, �ٴ������¼ C, ���ű� D" & vbNewLine & _
        " Where a.Id = b.��Դid and b.id=c.����id And a.����id = d.Id And a.�Ű෽ʽ =[3] And Nvl(a.�Ƿ�ɾ��, 0) = 0" & vbNewLine & _
        "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
        "       And (Nvl([2], 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = [2])))" & vbNewLine & _
        "       And b.����id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ��������Ƿ������Ч�İ���", lng����ID, lngUserID, byt�Ű෽ʽ)
    ExistsPlanOnVisitTable = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ܳ����ID(ByVal intYear As Integer, ByVal intMonth As Integer, _
    ByVal intWeek As Integer) As Long
    '���������ܻ�ȡ�ܰ��ŵĳ����ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    strSQL = "Select b.ID" & vbNewLine & _
            " From �ٴ������ B" & vbNewLine & _
            " Where Nvl(�Ű෽ʽ, 0) = 2 And ��� = [1] And �·� = [2] And ���� = [3] And Nvl(վ��,'-') = Nvl([4],'-')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���������ܻ�ȡ�ܰ��ŵĳ����ID", intYear, intMonth, intWeek, gstrNodeNo)
    If Not rsTemp.EOF Then
        Get�ܳ����ID = Val(Nvl(rsTemp!id))
    End If
End Function

Public Function GetNewPlanInfo(ByVal frmParent As Form, ByVal strPrivs As String, ByVal blnMonth As Boolean, _
    ByRef strCurrentPlanKey As String, ByRef blnDeletePlan As Boolean) As Collection
    '��ȡ�ͼ����һ����������Ϣ
    '��Σ�
    '   blnMonth - �Ƿ��³����
    '���Σ�
    '   strCurrentPlanKey - �����Keyֵ�����ڶ�λ
    '   blnDeletePlan - �Ƿ�ɾ����δ���ڹҺ�ԤԼ�Ĺ̶����ﰲ��
    '���أ��³������Ϣ��Array(���,�·�,����,��ʼ����,��������)
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim varDateRange As Variant, dtStart As Date, dtEnd As Date
    Dim cllPlan As New Collection 'Array(���,�·�,����,��ʼ����,��������)
    Dim intElseYear As Integer, intElseMonth As Integer, intElseWeek As Integer
    Dim dtStartTemp As Date, dtEndTemp As Date
    Dim strDelInfo As String, strNotDelInfo As String, strInfo As String
    
    Err = 0: On Error GoTo errHandler
    strCurrentPlanKey = "": blnDeletePlan = False
    
    If GetNextPlanDate(frmParent, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, _
        IIf(HavePrivs(strPrivs, "���п���"), 0, UserInfo.id)) = False Then
        MsgBox "ȷ����һ������������ʱ���������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    'XX�³����ڵ㣺K2_���_�·�
    'XX�ܳ����ڵ㣺K3_���_�·�_����
    If blnMonth Then
        strCurrentPlanKey = "K2_" & intYear & "_" & intMonth
    Else
        strCurrentPlanKey = "K3_" & intYear & "_" & intMonth & "_" & intWeek
    End If
    
    varDateRange = GetDateRange(intYear, intMonth, intWeek)
    dtStart = varDateRange(0): dtEnd = varDateRange(1)
    'Array(���,�·�,����,��ʼ����,��������)
    cllPlan.Add Array(intYear, intMonth, intWeek, dtStart, dtEnd)
    
    '���ܿ��µ��ܳ�����ͬ������
    If blnMonth = False Then
        dtStartTemp = dtStart: dtEndTemp = dtEnd
        If IsDoubleMonthWeekPlan(intElseYear, intElseMonth, intElseWeek, dtStartTemp, dtEndTemp) Then
            dtStart = dtStartTemp: dtEnd = dtEndTemp
            
            varDateRange = GetDateRange(intElseYear, intElseMonth, intElseWeek)
            'Array(���,�·�,����,��ʼ����,��������)
            cllPlan.Add Array(intElseYear, intElseMonth, intElseWeek, varDateRange(0), varDateRange(1))
        Else
            intElseYear = 0: intElseMonth = 0: intElseWeek = 0
        End If
    End If
    
    If CheckExistsFixedToOthers(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "���п���"), 0, UserInfo.id), dtStart, strDelInfo, strNotDelInfo) Then
        strInfo = "��ʾ��" & vbCrLf & _
                  "     ��ǰ�ɰ�" & IIf(blnMonth, "��", "��") & "�Ű�Ĳ��ֺ�Դ��ʱ��(" & Format(dtStart, "yyyy-mm-dd") & ")�Ժ���ڹ̶����ﰲ�š�"
        If strDelInfo <> "" Then
            strInfo = strInfo & vbCrLf & "���º�Դ��ɾ���ⲿ�ֳ��ﰲ�ţ�Ȼ��" & IIf(blnMonth, "��", "��") & "�Űࣺ" & vbCrLf & strDelInfo & vbCrLf & _
                "�Ƿ�ɾ���ⲿ�ֳ��ﰲ�ţ��Ա���ɰ�" & IIf(blnMonth, "��", "��") & "�Űࣿ"
            If strNotDelInfo <> "" Then
                strInfo = strInfo & vbCrLf & vbCrLf & _
                    "��" & vbCrLf & _
                    "     ���º�Դ��Ϊ�ⲿ�ֳ��ﰲ���еĲ��������ڹҺ�ԤԼ�����ܰ�" & IIf(blnMonth, "��", "��") & "�Űࣺ" & vbCrLf & strNotDelInfo
            End If
            blnDeletePlan = MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        Else
            strInfo = strInfo & vbCrLf & "���º�Դ��Ϊ�ⲿ�ֳ��ﰲ���еĲ��������ڹҺ�ԤԼ�����ܰ�" & IIf(blnMonth, "��", "��") & "�Űࣺ" & vbCrLf & strNotDelInfo & vbCrLf & _
                "�Ƿ������"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    '����Ƿ��а���/���Ű����Ч��Դ
    If cllPlan.Count = 2 Then
        'Array(���,�·�,����,��ʼ����,��������)
        dtStart = cllPlan(2)(3): dtEnd = cllPlan(2)(4)
        If CheckIsHavePlan(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "���п���"), 0, UserInfo.id), _
            dtStart, dtEnd, blnDeletePlan) = False Then
            cllPlan.Remove 2
        End If
    End If
    'Array(���,�·�,����,��ʼ����,��������)
    dtStart = cllPlan(1)(3): dtEnd = cllPlan(1)(4)
    If CheckIsHavePlan(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "���п���"), 0, UserInfo.id), _
        dtStart, dtEnd, blnDeletePlan) = False Then
        cllPlan.Remove 1
    End If
    
    If cllPlan.Count = 0 Then
        MsgBox "��ǰ�ް�" & IIf(blnMonth, "��", "��") & "�Ű����Ч��Դ�����ȵ�����������>�ٴ���Դ��������ӣ�", vbInformation, gstrSysName
        Exit Function
    End If
    Set GetNewPlanInfo = cllPlan
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

