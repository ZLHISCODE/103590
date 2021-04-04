Attribute VB_Name = "mdlBlood"
Option Explicit
Public gobjCardSquare As Object 'һ��ͨ����
Public gobjPublicExpense As Object '���ù�������
Public gobjRegister As Object          'ע����Ȩ����zlRegister

Private mstrSQL As String
Public Enum COLOR
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    
    ԭʼ���� = 0
    ������¼ = &HFF
    ͣ����Ŀ = &H8000000C
    ������Ŀ = 0
    
    ����ģ��ɫ = &HC00000
    
    
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
    
End Enum

Public Enum Enum_Inside_Program
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    pҽ�����ѹ��� = 1257
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    P�°滤ʿվ = 1265
    p��Ѫ��˹��� = 1268
    p��Ѫ��Ӧ���� = 1938
    pѪҺ���յǼ� = 1910
End Enum

Public Function GetObjectRegister() As Boolean
'����ע����Ȩ����zlRegister
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
    
        If gobjRegister Is Nothing Then
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If gobjRegister Is Nothing Then
                MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    GetObjectRegister = True
End Function

Public Function InitObjPublicExpense(ByVal lngSys As Long) As Boolean
    If gobjPublicExpense Is Nothing Then
        On Error Resume Next
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not gobjPublicExpense Is Nothing Then
            Call gobjPublicExpense.zlInitCommon(lngSys, gcnOracle, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    InitObjPublicExpense = Not gobjPublicExpense Is Nothing
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjCardSquare Is Nothing Then
        Set gobjCardSquare = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjCardSquare.zlInitComponents(frmMain, lngModule, lngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjCardSquare Is Nothing Then Exit Sub
    If Not gobjCardSquare Is Nothing Then
         Set gobjCardSquare = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
End Sub



Public Function GetDeptList(ByVal str�������� As String, Optional ByVal int������� As Integer = -1, Optional ByVal blnShowAll As Boolean = True) As ADODB.Recordset
        '******************************************************************************************************************
    '���ܣ���ȡ����ѡ��
    '������
    '���أ����ؼ�¼��
    '******************************************************************************************************************
    Dim bytService(3) As Byte
    
    Select Case int�������
    Case -1         '���ж�
        bytService(0) = 0
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    Case 0          '�������ڲ���
        bytService(0) = 0
        bytService(1) = 0
        bytService(2) = 0
        bytService(3) = 0
    Case 1          '���������ﲡ��
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 9
        bytService(3) = 3
    Case 2          '������סԺ����
        bytService(0) = 9
        bytService(1) = 9
        bytService(2) = 2
        bytService(3) = 3
    Case 3          '�����������סԺ����
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    End Select
        
    If blnShowAll Then '���в���
        mstrSQL = "SELECT Distinct A.����,A.����, A.ID,A.���� FROM ���ű� A,��������˵�� B WHERE B.������� In ([2],[3],[4],[5]) And (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������=[1] ORDER BY A.����"
        Set GetDeptList = gobjDatabase.OpenSQLRecord(mstrSQL, "��ȡ�����б�", str��������, bytService(0), bytService(1), bytService(2), bytService(3))
    Else
        mstrSQL = _
            " Select Distinct a.����, a.����, a.Id, a.����, c.ȱʡ" & vbNewLine & _
            " From ���ű� a, ��������˵�� b, ������Ա c" & vbNewLine & _
            " Where (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And a.Id = b.����id And b.�������� = [1] And" & vbNewLine & _
            "      b.������� In ([3],[4],[5],[6]) And a.Id = c.����id And c.��Աid = [2]" & vbNewLine & _
            " Order By a.���� || '-' || a.����"
        Set GetDeptList = gobjDatabase.OpenSQLRecord(mstrSQL, "��ȡ�����б�", str��������, UserInfo.id, bytService(0), bytService(1), bytService(2), bytService(3))
    End If

End Function

Public Function GetPatientOtherInfo(ByVal lng����ID As Long, ByVal str��Ϣ�� As String) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡ������Ϣ�ӱ�����
    '������
    '���أ����ؼ�¼��
    '******************************************************************************************************************

    mstrSQL = "Select ����id,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����id=[1] And ��Ϣ��=[2]"
    Set GetPatientOtherInfo = gobjDatabase.OpenSQLRecord(mstrSQL, "������Ϣ�ӱ�", lng����ID, str��Ϣ��)
End Function


Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��������е�һЩ���ܣ���ͼ�꣬��׼��ť��״̬����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
    
    Select Case Control.id
        Case conMenu_View_ToolBar_Button '������
        
            For lngLoop = 2 To frmMain.cbsMain.Count
                frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
            Next
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_ToolBar_Text '��ť����
        
            For lngLoop = 2 To frmMain.cbsMain.Count
                For Each objControl In frmMain.cbsMain(lngLoop).Controls
                    If objControl.Type = xtpControlButton Then
                        objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                    End If
                Next
            Next
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_ToolBar_Size '��ͼ��
        
            frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_StatusBar '״̬��
        
            frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
            frmMain.cbsMain.RecalcLayout
                
    End Select
    CommandBarExecutePublic = True
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '����:��ȡ����ʱ��
    '����:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(gobjDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(gobjDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case Else
        If strMode = Val(strMode) Then
            If bytFlag = 1 Then
                GetDateTime = Format(DateAdd("d", -Val(strMode), CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
            End If
        End If
    End Select
    
End Function

Public Function GetDepartPeople(ByVal str�������� As String, Optional ByVal int������� As Integer = -1, Optional ByVal blnShowAll As Boolean = True)
Dim bytService(3) As Byte
    
    Select Case int�������
    Case -1         '���ж�
        bytService(0) = 0
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    Case 0          '�������ڲ���
        bytService(0) = 0
        bytService(1) = 0
        bytService(2) = 0
        bytService(3) = 0
    Case 1          '���������ﲡ��
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 9
        bytService(3) = 9
    Case 2          '������סԺ����
        bytService(0) = 9
        bytService(1) = 9
        bytService(2) = 2
        bytService(3) = 9
    Case 3          '�����������סԺ����
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    End Select
    
    If blnShowAll Then
        mstrSQL = " Select Distinct d.���� " & _
                  " From ��������˵�� a, ���ű� b, ������Ա c, ��Ա�� d " & _
                  " Where a.�������� = [1] And a.����id = b.Id And c.����id = b.Id And c.��Աid = d.Id and a.������� in([2],[3],[4],[5])"
        Set GetDepartPeople = gobjDatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ա��Ϣ", str��������, bytService(0), bytService(1), bytService(2), bytService(3))
    Else
        mstrSQL = "select ���� from ��Ա�� where id=[1]"
        Set GetDepartPeople = gobjDatabase.OpenSQLRecord(mstrSQL, "��ȡ������Ա��Ϣ", UserInfo.id)
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ҽ���������
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlBloodInstantRptPrint(ByVal objFrm As Object, ByVal lngActiveID As Long) As Boolean
'���ܣ�����ʿվ��ҽ������վ����(��Ѫִ�е���ӡ)
'������ objFrm--����������
'           lngActiveID--ҽ��ID
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = _
        " Select  c.��������, c.ִ�з���,b.ҽ��״̬" & vbNewLine & _
        " From ������ĿĿ¼ c, ����ҽ����¼ a, ����ҽ����¼ b" & vbNewLine & _
        " Where c.Id = a.������Ŀid And a.���id = b.Id And a.������� = 'E' And b.Id = [1] And b.������� = 'K'"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "zlBloodInstantRptPrint", lngActiveID)
    If rsTmp.EOF Then
        MsgBox "ѡ�е�ҽ��������Ѫҽ������ѡ����Ѫҽ����", vbInformation, gstrSysName
        Exit Function
    End If
    If Not (Val("" & rsTmp!��������) = 8 And Val("" & rsTmp!ִ�з���) = 1) Then
        MsgBox "ѡ�е�ҽ��������Ѫ�����Ѫҽ������ѡ����Ѫҽ����", vbInformation, gstrSysName
        Exit Function
    End If
    zlBloodInstantRptPrint = frmBloodInstantRptPrint.ShowMe(objFrm, lngActiveID)
End Function

Public Function zlAdviceOperation(ByVal lngMoudle As Long, ByVal lngҽ��ID As Long, ByVal intOperation As Enum_Advice, Optional ByVal blnMoved As Boolean = False, _
        Optional ByRef strErrInfo As String = "") As Boolean
'���ܣ�ҽ���������ýӿڣ��¿���ɾ�������͡�����ʱ�˷����ĵ��������ҽ�����������е��ã��޸ġ�У�ԡ�����Ϊ����У���飬��������֮ǰ��
'���:
'       lngMoudle:����ģ���
'       lngҽ��ID:ѪҺҽ����ҽ��ID
'       intOperation:ҽ����������(ö��),�����¿����޸ġ�ɾ����У�ԡ����ϡ����͡�����
'       blnMoved:������ʷ�����Ƿ�ת��
'���Σ�
'       strErrInfo���ӿڷ���FALSEʱ����Ϣ
'���أ��ɹ�=TRUE��ʧ��=False
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    'ҽ�����ݱ���
    Dim int������Դ As Integer, lng����ID As Long, lng��ҳid As Long, lngִ�п���ID As Long, lng���ID As Long
    Dim int���״̬ As Long, str��鷽�� As String, int�������� As Integer, intִ�з��� As Integer, intҽ��״̬ As Integer
    Dim bln��Ѫ As Boolean
    
    On Error GoTo ErrHand
    If blnMoved = True Then
        strErrInfo = "���˵������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�"
        Exit Function
    End If
    strSQL = _
        " Select a.id ���ID,b.������Դ, b.����id, b.��ҳid, b.ִ�п���id, b.���״̬, b.��鷽��, c.��������, c.ִ�з���,B.ҽ��״̬" & vbNewLine & _
        " From ������ĿĿ¼ c, ����ҽ����¼ a, ����ҽ����¼ b" & vbNewLine & _
        " Where c.Id = a.������Ŀid And a.���id = b.Id And a.������� = 'E' And b.Id = [1] And b.������� = 'K'"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "zlAdviceOperation", lngҽ��ID)
    'ѪҺҽ���϶���õ����ݣ��鲻���������˳�
    If rsTmp.EOF Then
        zlAdviceOperation = True
        Exit Function
    End If
    
    lng���ID = Val("" & rsTmp!���ID)
    int������Դ = Val("" & rsTmp!������Դ)
    lng����ID = Val("" & rsTmp!����id)
    lng��ҳid = Val("" & rsTmp!��ҳid)
    lngִ�п���ID = Val("" & rsTmp!ִ�п���ID)
    int���״̬ = Val("" & rsTmp!���״̬)
    str��鷽�� = "" & rsTmp!��鷽��
    int�������� = Val("" & rsTmp!��������)
    intִ�з��� = Val("" & rsTmp!ִ�з���)
    intҽ��״̬ = Val("" & rsTmp!ҽ��״̬)
    If str��鷽�� = "" Then
        If int�������� = "8" And intִ�з��� = 1 Then
            bln��Ѫ = True
        End If
    Else
        bln��Ѫ = Val(str��鷽��) = 1
    End If
    Select Case intOperation
        Case Advice_�¿�
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If bln��Ѫ = True And gblnҽ�����ͺ�Ѫ = False Then
                strSQL = "Zl_ѪҺҽ����¼_Insert(" & lngҽ��ID & "," & lng����ID & "," & IIf(lng��ҳid = 0, "NULL", lng��ҳid) & "," & int������Դ & "," & lngִ�п���ID & "," & 2 & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Insert")
            End If
        Case Advice_�޸�
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            
            If int���״̬ = 5 Or int���״̬ = 2 Then
                strErrInfo = "��ҽ��Ŀǰ��Ѫ���Ѿ����գ��������ҽ�����в�����"
                Exit Function
            End If
            
            If bln��Ѫ = True And intҽ��״̬ = 1 Then  'ҽ�������¿�״̬�����޸�ѪҺ��Ѫ��¼��Ϣ��
                '����Ƿ������Ѫ��¼
                strSQL = "select id From  ѪҺ��Ѫ��¼ where ����ID=[1]"
                Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "zlAdviceOperation", lngҽ��ID)
                If Not rsData.EOF Then
                    strSQL = "Zl_ѪҺҽ����¼_Insert(" & lngҽ��ID & "," & lng����ID & "," & IIf(lng��ҳid = 0, "NULL", lng��ҳid) & "," & int������Դ & "," & lngִ�п���ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Insert")
                End If
            End If
        Case Advice_ɾ��
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int���״̬ = 5 Or int���״̬ = 2 Then
                strErrInfo = "��ҽ��Ŀǰ��Ѫ���Ѿ����գ��������ҽ�����в�����"
                Exit Function
            End If
            strSQL = "Zl_ѪҺҽ����¼_Delete(" & lngҽ��ID & "," & IIf(bln��Ѫ = False, 1, 2) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Delete")
        Case Advice_У��
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int���״̬ = 5 Or int���״̬ = 2 Then
                zlAdviceOperation = True
                Exit Function
            End If
            'ҽ��У��ֱ�ӷ��͵����
            If intҽ��״̬ = 8 Then
                If bln��Ѫ = False Or (bln��Ѫ = True And gblnҽ�����ͺ�Ѫ = True) Then
                    strSQL = "Zl_ѪҺҽ����¼_Insert(" & lngҽ��ID & "," & lng����ID & "," & IIf(lng��ҳid = 0, "NULL", lng��ҳid) & "," & int������Դ & "," & lngִ�п���ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Insert")
                End If
            End If
        Case Advice_����
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int���״̬ = 5 Or int���״̬ = 2 Then
                strErrInfo = "��ҽ��Ŀǰ��Ѫ���Ѿ����գ��������ҽ�����в�����"
                Exit Function
            End If
            '���ﲡ������ɾ����Ѫ��Ϣ��סԺ���˱�Ѫҽ������ɾ������Ѫҽ����������ɾ���ͻ���ɾ�������(���ݲ�����gblnҽ�����ͺ�Ѫ����)
            strSQL = "Zl_ѪҺҽ����¼_Delete(" & lngҽ��ID & "," & IIf(bln��Ѫ = False, 1, 2) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Delete")
        Case Advice_�������� 'סԺ����ҽ�����Ͽ��Ի���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            '��Ҫ��סԺ���˵���Ѫҽ������Ϊ��Ѫҽ���ķ��ͺ�Ų�����Ѫ��Ϣ
            If int������Դ = 2 Then
                If bln��Ѫ = True And gblnҽ�����ͺ�Ѫ = False Then
                    strSQL = "Zl_ѪҺҽ����¼_Insert(" & lngҽ��ID & "," & lng����ID & "," & IIf(lng��ҳid = 0, "NULL", lng��ҳid) & "," & int������Դ & "," & lngִ�п���ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Insert")
                End If
            End If
        Case Advice_����
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            strSQL = "Zl_ѪҺҽ����¼_Insert(" & lngҽ��ID & "," & lng����ID & "," & IIf(lng��ҳid = 0, "NULL", lng��ҳid) & "," & int������Դ & "," & lngִ�п���ID & "," & IIf(bln��Ѫ = True, 2, 1) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Insert")
        Case Advice_����
            '�ϵ���Ѫҽ���������κδ���
            If bln��Ѫ = True And str��鷽�� = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int���״̬ = 5 Or int���״̬ = 2 Then
                strErrInfo = "��ҽ��Ŀǰ��Ѫ���Ѿ����գ��������ҽ�����в�����"
                Exit Function
            End If
            If bln��Ѫ = False Or (bln��Ѫ = True And gblnҽ�����ͺ�Ѫ = True) Then
                strSQL = "Zl_ѪҺҽ����¼_Delete(" & lngҽ��ID & "," & IIf(bln��Ѫ = False, 1, 2) & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_ѪҺҽ����¼_Delete")
            End If
    End Select
    zlAdviceOperation = True
    Exit Function
ErrHand:
    If gcnOracle.Errors.Count <> 0 Then
        strErrInfo = gcnOracle.Errors(0).Description
        If InStr(UCase(strErrInfo), "[ZLSOFT]") > 0 Then
            strErrInfo = Split(strErrInfo, "[ZLSOFT]")(1)
        End If
    Else
        strErrInfo = Err.Description
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'ҽ��ִ�����
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ItemCanCancel(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng��ID As Long, str������� As String, _
    ByVal bln����ִ�� As Boolean, ByVal blnMove As Boolean, ByVal byt��Դ As Byte) As Boolean
'���ܣ��ж�ָ����Ŀ�Ƿ����ȡ��ִ��
'������byt��Դ=1:���2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If gbytBillOpt = 0 Then ItemCanCancel = True: Exit Function
    
    On Error GoTo errH
    
    If bln����ִ�� Then
        strSQL = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]"
    Else
        strSQL = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[2]" & _
            " And ҽ��ID IN(Select ID From ����ҽ����¼ Where (ID=[3] Or ���ID=[3]) And �������=[4])"
    End If
    If blnMove Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "ItemCanCancel", lngҽ��ID, lng���ͺ�, lng��ID, str�������)
    
    Do While Not rsTmp.EOF
        '�������ſ��˽��ʽ��Ϊ0�ģ�����ķ��õǼ�
        If HaveBilling(rsTmp!NO, True, "", IIf(bln����ִ��, lngҽ��ID, 0), byt��Դ) <> 0 Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("����Ŀ�����Ѿ����ʵķ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    MsgBox "����Ŀ�����Ѿ����ʵķ���,�������ܼ�����", vbExclamation, gstrSysName
                    Exit Function
            End Select
        End If
        rsTmp.MoveNext
    Loop
    ItemCanCancel = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, ByVal blnALL As Boolean, _
     ByVal strTime As String, ByVal lngҽ��ID As Long, ByVal byt��Դ As Byte) As Integer
'���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
'������strNO=���ʵ��ݺ�,�������ＰסԺ
'      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�(����ʱ)
'      byt��Դ=1:���2-סԺ
'���أ�0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    Dim strTab As String
    
    On Error GoTo errH
    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
        
    '��δ���ϵķ�����
    strSQL = _
        " Select ��� From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼����=2" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    
    '��ÿ�еĽ������
    strSQL = _
        "Select Nvl(�۸񸸺�,���) as ���,Sum(Nvl(���ʽ��,0)) as ���ʽ��" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼���� IN(2,12)" & _
        IIf(Not blnALL, " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
        IIf(lngҽ��ID <> 0, " And ҽ�����+0=[3]", "") & _
        " Group by Nvl(�۸񸸺�,���)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "HaveBilling", strNO, CDate(IIf(strTime = "", "1990-01-01", strTime)), lngҽ��ID)
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '��������
        rsTmp.Filter = "���ʽ��<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '�޽�����
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 'ȫ�����ѽ���
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '�������ѽ���
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat����ʱ�� As Date, Optional ByRef strҽ��IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
'������int������Դ=1-����,2-סԺ
'      str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
'      int��ʽ=0-����Ƿ����δ�շѼ�¼
'              1-����Ƿ�������շѼ�¼
'      int�������=1=סԺ���͵��������
'      ���أ�strҽ��IDs=��ҽ������ص�ҽ��ID,NOs=ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTab As String
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    ItemHaveCash = True
    strҽ��IDs = ""
    strNOs = ""
    
    '��Ӧ�ķ������Ƿ����δ�շ�[��������]������
    '���嵥ֻ��ʾ���շѲ�ͬ��
    '1.�����ҽ������(���Ӽ�¼���ʵ���������Ϊ���ܲ��շѵ�����ʵ�)
    '2.���ʻ���Ҳ��ʾΪδ��(�嵥��Ҫ���Գ���ִ�к����)
    '3.��NO��Ӧ�����ҽ���ķ��ü��(�嵥�ǰ���ʾ��ҽ��ID)
    strSQL = _
        " Select A.��¼״̬,Nvl(B.���ID,B.ID) as ҽ��ID,B.�������,A.ִ��״̬,A.NO" & IIf(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(A.����״̬,0) as ����״̬") & _
        " From " & strTab & " A,����ҽ����¼ B" & _
        " Where A.NO=[4] And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And A.��¼����=[5]" & IIf(bln����ִ��, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.��¼״̬,Nvl(C.���ID,C.ID) as ҽ��ID,C.�������,B.ִ��״̬,A.NO" & IIf(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(b.����״̬,0) as ����״̬") & _
        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
        " Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0" & IIf(bln����ִ��, " And A.ҽ��ID=[2]", _
            " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
        " And A.���ͺ�=[3] And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����=[5]"
    If blnMove Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf gobjDatabase.DateMoved(dat����ʱ��) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
    If Not rsTmp.EOF Then
        If int��ʽ = 0 Then
            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ����״̬=1"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬=0"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
                    strҽ��IDs = strҽ��IDs & "," & rsTmp!ҽ��ID
                End If
                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNOs = strNOs & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
            strNOs = Mid(strNOs, 2)
            strҽ��IDs = Mid(strҽ��IDs, 2)
        ElseIf int��ʽ = 1 Then
            rsTmp.Filter = "ҽ��ID=" & IIf(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬<>1 And ����״̬<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int��ʽ = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal str��ID As String, ByVal strҽ��ID As String, ByVal str���ͺ� As String, _
    str��� As String, str����� As String, ByVal bln����ִ�� As Boolean, ByVal byt��Դ As Byte) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      byt��Դ��1:���2-סԺ
'���أ�str���,str�����=���ڱ�����ʾ
'˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strTab As String
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
     
    strTab = IIf(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    
    If bln����ִ�� Then
        strSQL = _
            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    Else
        strSQL = _
            " Select /*+ RULE */ B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From ����ҽ����¼" & _
            "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From ����ҽ����¼" & _
                "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str��ID, strҽ��ID, str���ͺ�)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳid As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
'���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(strPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
        Exit Function
    End If
    On Error GoTo errH
    strSQL = "Select NVL(B.����,A.����) ����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+) And X.����(+) = 2" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳid)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, "��Ժ����ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If lngModual = pҽ�����ѹ��� Or lngModual = pסԺҽ������ Or lngModual = pסԺҽ���´� Then
            '68081�������Ժ���˴���ҽ������
            strMsg = """" & rsTmp!���� & """�Ѿ���Ժ(��Ԥ��Ժ)�����ܶԸò��˵�ҽ�����з��͡������ջء�ִ�С����ˡ�"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function IsBloodMessageDone(ByVal intMode As Integer, ByVal lng����ID As Long, ByVal lng����id As Long, _
                                    ByVal int�Ķ����� As Integer, ByVal lng�Ķ�����id As Long) As Boolean
'������intMode  1-Ѫ���Ƿ��Ѿ����գ�2-Ѫ���Ƿ���д��Ѫ��Ӧ
'���ܣ���ѯҽ��վ��Ѫ�������Ϣ�Ƿ�����˺�������
    Dim rsTmp As New ADODB.Recordset, rsMsg As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lng�շ�id As Long, str�շ�ids As String
    Dim arr() As String
    Dim blnTrans As Boolean, arrSQL As Variant
    On Error GoTo errH
    arrSQL = Array()
    Select Case intMode
        Case 1
            strSQL = "select id,ҵ���ʶ from ҵ����Ϣ�嵥 where ���ͱ��� = [1] and ����ID = [2] and ����id = [3] and �Ƿ����� = 0 "
            Set rsMsg = gobjDatabase.OpenSQLRecord(strSQL, "��Ϣ״̬", "ZLHIS_BLOOD_007", lng����ID, lng����id)
            Do While Not rsMsg.EOF
                arr = Split(rsMsg!ҵ���ʶ, ":")
                If UBound(arr) > 0 Then
                    str�շ�ids = str�շ�ids & ":" & Val(arr(2))
                End If
                rsMsg.MoveNext
            Loop
            strSQL = "select /*+ CARDINALITY(c,10) */" & vbNewLine & _
                    "       a.id, b.��¼״̬ from ѪҺ�շ���¼ a ,ѪҺ��Ѫ���� b,table(f_str2list([1],':')) c" & vbNewLine & _
                    "       where a.id = b.�շ�id and b.Ѫ����� = a.Ѫ����� and a.id = c.column_value"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "����״̬", Mid(str�շ�ids, 2, Len(str�շ�ids)))
            
                rsTmp.Filter = "��¼״̬ = '1'"
            If rsTmp.RecordCount = 0 Then
                strSQL = "zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_007'," _
                                                & int�Ķ����� & ",'" & UserInfo.���� & "'," & lng�Ķ�����id & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                IsBloodMessageDone = True
                Else
                '����¼״̬Ϊ1����Ϣ���շ�id�Ӵ����޳�
                rsTmp.MoveFirst
                str�շ�ids = str�շ�ids & ":"
                Do While Not rsTmp.EOF
                    str�շ�ids = Replace(str�շ�ids, ":" & rsTmp!id & ":", ":")
                    rsTmp.MoveNext
                Loop
                '��������ѪҺ���յ���Ϣ���շ�id���д�ŵ�Ϊ��Ϊ��¼״̬��Ϊ1�����Ѫ�����շ�id�����ⲿ�ְ�����Ϊ�Ѷ�
                rsMsg.MoveFirst
                Do While Not rsMsg.EOF
                    rsTmp.MoveFirst
                        If InStr(str�շ�ids, Mid(rsMsg("ҵ���ʶ"), InStr(rsMsg("ҵ���ʶ"), ":"))) > 0 Then
                            strSQL = "zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_007'," _
                                                    & int�Ķ����� & ",'" & UserInfo.���� & "'," & lng�Ķ�����id & ",null," & rsMsg("ID") & " ,null)"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                End If
                    rsMsg.MoveNext
                Loop
                IsBloodMessageDone = False
            End If
        Case 2
            strSQL = "select id,ҵ���ʶ from ҵ����Ϣ�嵥 where ���ͱ��� = [1] and ����ID = [2] and ����id = [3] and �Ƿ����� = 0 "
            Set rsMsg = gobjDatabase.OpenSQLRecord(strSQL, "��Ϣ״̬", "ZLHIS_BLOOD_006", lng����ID, lng����id)
            Do While Not rsMsg.EOF
                arr = Split(rsMsg!ҵ���ʶ, ":")
                If UBound(arr) > 0 Then
                    str�շ�ids = str�շ�ids & ":" & Val(arr(1))
                End If
                rsMsg.MoveNext
            Loop
            If str�շ�ids = "" Then IsBloodMessageDone = False: Exit Function
            strSQL = "SELECT /*+ CARDINALITY(b,10) */ a.�շ�id, a.������Ѫ��Ӧ FROM ��Ѫ��Ӧ��¼ a, TABLE(f_Str2list([1], ':')) b " _
                    & "WHERE a.�շ�id = b.Column_Value"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��Ѫ��Ӧ��¼", Mid(str�շ�ids, 2, Len(str�շ�ids)))
            
            If rsTmp.RecordCount = rsMsg.RecordCount Then           'ÿ��δ����Ϣ���ж�Ӧ����Ѫ��Ӧ��¼����ʾ������д��������δ����Ϣ��Ϊ�Ѷ�
                strSQL = "zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_006'," _
                                                & int�Ķ����� & ",'" & UserInfo.���� & "'," & lng�Ķ�����id & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                IsBloodMessageDone = True
            Else            '����������Ѫ��Ӧ��¼��Ѫ������Ӧ����Ϣ��Ϊ�Ѷ�
                If rsTmp.RecordCount <> 0 Then
                    rsMsg.MoveFirst
                    Do While Not rsMsg.EOF
                        rsTmp.MoveFirst
                        Do While Not rsTmp.EOF
                            If InStr(rsMsg("ҵ���ʶ") & ":", ":" & rsTmp("�շ�id") & ":") > 0 Then
                                strSQL = "zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����id & ",'ZLHIS_BLOOD_007'," _
                                                    & int�Ķ����� & ",'" & UserInfo.���� & "'," & lng�Ķ�����id & ",null," & rsMsg("ID") & " ,null)"
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = strSQL
                                Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsMsg.MoveNext
                    Loop
                End If
                IsBloodMessageDone = False
            End If
    End Select
    If UBound(arrSQL) < 0 Then Exit Function
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), "ѪҺ�����Ϣ����")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Function
errH:
    IsBloodMessageDone = False
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function GetReactionTips(lng�ⷿid As Long) As Recordset
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "select x.��Ϣ����,x.ҵ���ʶ �շ�id,x.����id,x.����id,x.������Դ from ҵ����Ϣ�嵥 x, ҵ����Ϣ���Ѳ��� b" & vbNewLine & _
                "where x.���ͱ��� = 'ZLHIS_BLOOD_008' and x.id = b.��Ϣid and x.�Ƿ����� = 0 "
    If lng�ⷿid > 0 Then
        strSQL = strSQL & " and b.����id = [1] "
    Else
        strSQL = strSQL & "and b.����id in (SELECT Distinct A.ID FROM ���ű� A,��������˵�� B" & vbNewLine & _
                        "WHERE B.������� In (9,1,2,3) And (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID  AND B.��������='Ѫ��')"
    End If
    strSQL = strSQL & " Order by x.�Ǽ�ʱ�� desc "
    Set rs = gobjDatabase.OpenSQLRecord(strSQL, "��Ѫ��Ӧ��Ϣ", lng�ⷿid)
    Set GetReactionTips = rs
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
