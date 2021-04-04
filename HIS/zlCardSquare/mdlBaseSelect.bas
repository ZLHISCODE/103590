Attribute VB_Name = "mdlBaseSelect"
Option Explicit
Public gbln����վ����� As Boolean

Public Function zl_GetFieldValue(ByVal rsTemp As ADODB.Recordset, _
    Optional ByVal strShowFields As String = "����,����", _
    Optional ByVal strShowSplit As String = "-") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ�ֶε����ֵ
    '���:rsTemp-��¼��
    '     strShowFields-��ʾ���ֶ�
    '     strShowSplit-��ʾ�ķ����
    '����:
    '����:�ɹ�,������ص��ֶ�ֵ
    '����:���˺�
    '����:2009-03-06 11:59:19
    '-----------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strValue As String, strLeft As String, strRight As String
    varData = Split(strShowFields, ",")
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.RecordCount = 0 Then Exit Function
    
    Select Case strShowSplit
    Case "[", "[]", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "�ۣ�", "��", "��"
        strLeft = "��": strRight = "��"
    Case "[]", "[", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "{}", "{", "}"
        strLeft = "{": strRight = "}"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case Else
        strLeft = "": strRight = strShowSplit
    End Select
    
    strValue = ""
    With rsTemp
        For i = 0 To UBound(varData) - 1
            strValue = strValue & strLeft & nvl(.Fields(varData(i))) & strRight
        Next
        strValue = strValue & nvl(.Fields(varData(UBound(varData))))
    End With
    zl_GetFieldValue = strValue
End Function
Public Function zl_��ȡ������Ϣ(ByVal strCaption As String, Optional bln��Ա�������� As Boolean = False, _
    Optional str�������� As String = "", Optional strAddRoomCaption As String = "") As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '     str��������-��������(�Զ��ŷ���,R,Z��)
    '     strAddRoomCaption-�Ƿ�������(ֻ������ʾȫ�����ݲ���ʾʱ)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-03-03 14:47:10
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If str�������� <> "" Then str�������� = "," & str�������� & ","
    
    If str�������� <> "" Then
        strSQL = "" & _
        "   SELECT DISTINCT a.id,NULL as �ϼ�ID,a.����, a.����,A.����,A.ĩ��" & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� " & _
        "       AND instr([2],','||b.����||',')>0 " & _
        "       AND a.id = c.����id " & zl_��ȡվ������(True, "a") & _
        "       AND  a.����ʱ�� >=to_date('3000-01-01','yyyy-mm-dd')" & _
                IIf(bln��Ա�������� = False, "", " and ID in (Select ����id From ������Ա where ��Աid=[1]) ") & _
        "   order by a.����"
    Else
        If bln��Ա�������� Or gbln����վ����� Then
            strSQL = "" & _
            "   Select ID, NULL as �ϼ�ID,����,����,����, ĩ�� " & _
            "   From ���ű�" & _
            "   Where (����ʱ�� Is NULL Or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01')  " & zl_��ȡվ������ & _
                    IIf(bln��Ա�������� = False, "", " and ID in (Select ����id From ������Ա where ��Աid=[1]) ") & _
            "  order by ����"
        Else
            gstrSQL = ""
            If strAddRoomCaption <> "" Then
                gstrSQL = "" & _
                "  Select  -1 id, -NULL �ϼ�id,'' ����, '" & strAddRoomCaption & "' ����,'" & zlCommFun.zlGetSymbol(strAddRoomCaption, 0) & "' ����,0  as ĩ��" & _
                "   From dual  UNION ALL " & vbCrLf
            End If
            strSQL = gstrSQL & _
            "   Select ID," & IIf(strAddRoomCaption <> "", "nvl(�ϼ�id,-1)", "�ϼ�ID") & " as �ϼ�ID,����,����,����,ĩ�� " & _
            "   From ���ű�" & _
            "   Where (����ʱ�� Is NULL Or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01')  " & _
            "   Start With �ϼ�ID Is NULL Connect By Prior ID=�ϼ�ID"
        End If
    End If
    Set zl_��ȡ������Ϣ = zlDatabase.OpenSQLRecord(strSQL, strCaption, UserInfo.id, str��������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zl_SetCtlBackColor(varCtls As Variant, Optional ByVal frmMain As Form = Nothing)
    '-----------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��ı���ɫ
    '���:objCtl-ָ���Ŀؼ�
    '     frmMain-������
    '����:
    '����:���˺�
    '����:2009-03-03 11:55:10
    '-----------------------------------------------------------------------------------------------------------
    Dim lngNotEnabledBackColor As Long
    Dim objCtl As Object, i As Integer
    
    If IsArray(varCtls) = False Then
        varCtls = Array(varCtls)
    End If
        
    For i = 0 To UBound(varCtls)
        Set objCtl = varCtls(i)
        
        Select Case UCase(TypeName(objCtl))
        Case "TEXTBOX", "COMBOBOX"
        Case "CHECKBOX": Exit Sub
        Case "LABEL": Exit Sub
        Case "DTPICKER": Exit Sub
        Case UCase("CommandButton"): Exit Sub
        Case UCase("ListView"): Exit Sub
        End Select
        
        '��ֹ״̬Ϊ���屳��ɫ,������Ϊ��ɫ
        lngNotEnabledBackColor = &H8000000A
        If Not frmMain Is Nothing Then lngNotEnabledBackColor = frmMain.BackColor
        If objCtl.Enabled Then
            objCtl.BackColor = &H80000005
        Else
            objCtl.BackColor = lngNotEnabledBackColor
        End If
    Next
End Sub

Public Function Select����ѡ����(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional blnָ����Ա�������� As Boolean = False, _
    Optional strSQL As String = "", _
    Optional lngID As Long = 0, _
    Optional strTittle As String = "����ѡ����", _
    Optional strNote As String = "", _
    Optional strNotFindMsg As String = "û�����������Ĳ���,����!", _
    Optional strShowField As String = "����,����", _
    Optional strShowSplit As String = "-", _
    Optional lng��ԱID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��������ѡ����
    '���:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     blnָ����Ա��������-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '     strShowField-��ʾ������ֶ�
    '     strShowSplit-��ʾ�������(strShowFieldΪ���ֶ����ϲŴ���)
    '����:lngID-���ز���ID(��Ҫ��������)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-03-06 12:18:31
    '-----------------------------------------------------------------------------------------------------------
 
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, lngH As Long, strFind As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim blnTree As Boolean
    
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strKey = GetMatchingSting(strSearch, False)
    
    blnTree = (strSearch = "" And str�������� = "" And blnָ����Ա�������� = False And strSQL = "" And Not (gstrNodeNo <> "-"))
    
    
    If strSQL <> "" Then
        gstrSQL = strSQL
    Else
    
        If str�������� = "" And blnָ����Ա�������� = False Then
            gstrSQL = "" & _
            "   Select   /*+ rule */ a.Id," & IIf(blnTree, "a.�ϼ�id", "-1*NULL ") & " as �ϼ�ID,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
            "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ�� " & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = "" & _
            "   Select  /*+ rule */  distinct a.Id," & IIf(blnTree, "a.�ϼ�id", "-1*NULL ") & " as �ϼ�ID,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
            "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��" & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c " & _
            IIf(str�������� = "", "", "       ,(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.����=J.column_value ") & _
            "         AND a.id = c.����id " & _
            IIf(blnָ����Ա�������� = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
        "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) " & zl_��ȡվ������(True, "a") & ""
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
           strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If blnTree Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
   ' MsgBox GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX
    '���궨λ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    If strSearch = "" And str�������� = "" And blnָ����Ա�������� = False And strSQL = "" Then
        '�����¼�
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 1, strTittle, False, "", strNote, False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", strNote, False, False, True, sngX, sngY, lngH, blnCancel, False, False, IIf(lng��ԱID = 0, UserInfo.id, lng��ԱID), str��������, strKey)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    Call zlControl.ControlSetFocus(objCtl, True)
    
    '������ص�ֵ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp)
        objCtl.Tag = Val(rsTemp!id)
        zlCommFun.PressKey vbKeyTab
    End Select
    lngID = Val(nvl(rsTemp!id))
    Select����ѡ���� = True
End Function


Public Function Select��Աѡ����(ByVal frmMain As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng����ID As Long = 0, _
    Optional lng��ԱID As Long = 0, _
    Optional bln��������Ա��ʾ As Boolean = False, _
    Optional strSearchKey As String = "", _
    Optional str��Ա���� As String = "", _
    Optional str����ְ�� As String = "", _
    Optional strרҵ����ְ�� As String = "", _
    Optional strTittle As String = "��Աѡ����", _
    Optional strNote As String = "��ѡ����ص���Ա", _
    Optional strNotFindMsg As String = "δ�ҵ�ָ������Ա,����!", _
    Optional strShowField As String = "����", _
    Optional strShowSplit As String = "-") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������Ա
    '���:frmMain-���õĸ�����
    '     objCtl-�ؼ�(Ŀǰֻ֧���ı���)
    '     strKey-����Ľ�ֵ
    '     lng����ID-�����Ϊ��,��������Ա,����, ��ָ�������µ���Ա
    '     str��Ա����: ��ҽ��,ҽ��1... ��ʽ
    '     str����ְ��strרҵ����ְ��: ��ְ��1,ְ��21... ��ʽ
    '����:lng��Աid-������ԱID
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, bytType As Byte, str��Ա����Table As String, strWhere As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmMain=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Err = 0: On Error GoTo Errhand:
    bytType = 0: strWhere = ""
    If str��Ա���� <> "" Then
        str��Ա����Table = "��Ա����˵�� Q1,(Select Column_Value From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) Q2" & vbCrLf
        strWhere = strWhere & " And ( A.ID=Q1.��ԱID and Q1.��Ա���� = Q2.Column_Value ) " & vbCrLf
    End If
    If str����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))  Where a.����ְ��=Column_Value) " & vbCrLf
    If strרҵ����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))  Where a.רҵ����ְ��=Column_Value) " & vbCrLf
    
    
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
        If lng����ID = 0 Then
            gstrSQL = "" & _
                "   Select /*+ Rule */  A.ID,A.���,A.����,A.����,A.����,A.�Ա�,A.����,A.��������,A.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� A " & str��Ա����Table & _
                "   Where (A.���� like [1] or A.��� like [1] or A.���� like Upper([1]) or A.���� like [1]) " & strWhere & zl_��ȡվ������(True, "A") & "" & _
                "       and (A.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                "   order by A.���"
        Else
            gstrSQL = "" & _
                "   Select   /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2]   " & strWhere & zl_��ȡվ������(True, "a") & _
                "       and (a.���� like [1] or a.��� like [1] or a.���� like Upper([1]) or a.���� like [1]) " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                "   order by ���"
        End If
     Else
        If lng����ID = 0 Then
            If bln��������Ա��ʾ Then
                gstrSQL = "" & _
                "   Select   /*+ rule */ id," & IIf(gstrNodeNo <> "-", "1 as ����ID,-1*NULL as �ϼ�ID", "Level as ����ID,�ϼ�id") & " ,����,����,0 ĩ��,'' as ����,'' as ����,''as �Ա�,''as ����, to_date(Null,'yyyy-mm-dd')  as ��������, '' as  �칫�ҵ绰 ,'' ִҵ���, '' ����ְ��,'' רҵ����ְ��" & _
                "   From ���ű� " & _
                "   where ����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') " & zl_��ȡվ������() & _
                    IIf(gstrNodeNo <> "-", "", "   Start with �ϼ�id is null connect by prior id=�ϼ�id ") & _
                "   union all " & _
                "   Select a.ID,999999 AS ����ID,b.����id as �ϼ�ID,a.���,a.����,1 as ĩ��,����,����,�Ա�,����,��������,�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ�� " & _
                "   From ��Ա�� a,������Ա b  " & str��Ա����Table & _
                "   Where a.id=b.��Աid and b.ȱʡ=1  " & strWhere & zl_��ȡվ������(True, "a") & _
                "         And (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                "   Order by ����ID,����"
                bytType = 2
            Else
                gstrSQL = "" & _
                    "   Select  /*+ rule */ A.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                    "   From ��Ա�� A " & str��Ա����Table & _
                    "   Where (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & strWhere & zl_��ȡվ������(True, "a") & _
                    "   order by a.���"
            End If
        Else
            gstrSQL = "" & _
                "   Select   /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2] " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & strWhere & zl_��ȡվ������(True, "a") & _
                "   order by a.���"
        End If
    End If
   
   
   '���궨λ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, strTittle, bytType = 2, strSearchKey, strNote, bytType = 2, False, Not (bytType = 2), sngX, sngY, lngH, blnCancel, False, False, strKey, lng����ID, str��Ա����, str����ְ��, strרҵ����ְ��)
    
    lng��ԱID = 0
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If bytType = 2 Then
        strShowField = "," & strShowField & ",M_��,"
        strShowField = Replace(strShowField, ",���,", ",����,")
        strShowField = Replace(strShowField, ",����,", ",����,")
        strShowField = Mid(strShowField, 2)
        strShowField = Replace(strShowField, ",M_��,", "")
    End If
    
    '������ص�ֵ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
        objCtl.Tag = Val(rsTemp!id)
        zlCommFun.PressKey vbKeyTab
    End Select
    lng��ԱID = Val(nvl(rsTemp!id))
    rsTemp.Close
    Select��Աѡ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Select��Ӧ��(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    Optional ByVal blnCheckִ��Ч�� As Boolean = False, _
    Optional ByVal blnҩƷ As Boolean, Optional ByVal bln���� As Boolean = False, _
    Optional ByVal bln�豸 As Boolean = False, Optional ByVal bln���� As Boolean = False, _
    Optional ByVal bln���� As Boolean = False, Optional blnOnlyName As Boolean = False, _
    Optional lng��Ӧ��ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��Ӧ��
    '���:objCtl-����ؼ�
    '    strKey-ѡ��Ӧ��
    '
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-25 11:26:43
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean, str���� As String, bytType As Byte
    strKey = Replace(strKey, "'", "")
    '����ȱʡֵ,�Բ���ϵͳ��ȷ����صĹ�Ӧ��
    If Not (bln���� Or bln�豸 Or bln���� Or blnҩƷ Or bln����) Then
        blnҩƷ = Int(glngSys / 100) = 1
        bln���� = Int(glngSys / 100) = 1
        bln���� = Int(glngSys / 100) = 4
        bln�豸 = Int(glngSys / 100) = 6
    End If
    str���� = IIf(blnҩƷ, "1", "_")
    str���� = str���� & IIf(bln����, "1", "_")
    str���� = str���� & IIf(bln�豸, "1", "_")
    str���� = str���� & IIf(bln����, "1", "_")
    str���� = str���� & IIf(bln����, "1", "_")
    
    Err = 0: On Error GoTo Errhand:
    
    '����:����λ����ʾ,1λ--ҩƷ��Ӧ�� 2λ--���ʹ�Ӧ�̡���3λ--�豸��Ӧ�̡���4λ--����,   5-�������� ÿλ��1�����ʾ,1��ʾΪtrue,0Ϊfalse,�Ժ�����ϵͳ�ӵ�6λ��ʼ
    If strKey = "" Then
        '��Ҫ˫�б���ʽ
        gstrSQL = "" & _
            "   Select id,�ϼ�ID,����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,to_char(ִ��Ч��,'yyyy-mm-dd') as ִ��Ч�� ,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
            "   From ��Ӧ�� " & _
            "   Where  (����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null)  " & _
            "           and (( ���� like [2]  And nvl(ĩ��,0)=1 " & zl_��ȡվ������ & ") or nvl(ĩ��,0)=0 ) " & _
            "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
            "   Order by level,ID"
        bytType = 2
    Else
        gstrSQL = "" & _
            "   Select id, ����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,to_char(ִ��Ч��,'yyyy-mm-dd') as ִ��Ч�� ,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
            "   From ��Ӧ�� " & _
            "   where   ĩ��=1 and ���� like [2]  " & zl_��ȡվ������ & "  " & _
            "           and  (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
            "           and (���� like [1] or ���� like [1] or ���� like upper([1]))  "
            bytType = 0
    End If
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
    End If
 
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim sngX As Single, sngY As Single, lngH As Long
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, "��Ӧ��ѡ��", IIf(bytType = 2, True, False), "", "��ѡ������豸�Ĺ�Ӧ��", IIf(bytType = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey, str����)
    If blnCancel Then
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "�����ڷ��������Ĺ�Ӧ��,����!"
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If blnCheckִ��Ч�� Then
        If Not IsNull(rsTemp!ִ��Ч��) Then
            If Format(zlDatabase.Currentdate, "yyyy-mm-dd") > nvl(rsTemp!ִ��Ч��) Then
                If MsgBox("��Ӧ�̵�ִ��Ч���Ѿ�����,�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Call zlControl.ControlSetFocus(objCtl)
                    Exit Function
                End If
            End If
        End If
    End If
    Call zlControl.ControlSetFocus(objCtl)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, nvl(rsTemp!����), "[" & nvl(rsTemp!����) & "] " & nvl(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
            Else
                .Text = IIf(blnOnlyName, nvl(rsTemp!����), "[" & nvl(rsTemp!����) & "] " & nvl(rsTemp!����))
            End If
        End With
    Else
        With rsTemp
            objCtl.Text = "[" & nvl(!����) & "] " & nvl(!����)
            objCtl.Tag = nvl(!id)
        End With
        zlCommFun.PressKey vbKeyTab
    End If
    lng��Ӧ��ID = Val(nvl(rsTemp!id))
    Select��Ӧ�� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String, _
    Optional blnվ�� As Boolean = False, Optional blnNotMsg As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnվ��-�Ƿ����վ������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  "
    End If
    gstrSQL = gstrSQL & strWhere & IIf(blnվ��, zl_��ȡվ������, "") & _
    "   order by ����"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If blnδ�ҵ����� Then
            If zlCommFun.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str����
                    End If
                End With
            Else
                If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
           If blnNotMsg = False Then ShowMsgbox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, nvl(rsTemp!����), nvl(rsTemp!����) & "-" & nvl(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!����)
            Else
                .Text = IIf(blnOnlyName, nvl(rsTemp!����), nvl(rsTemp!����) & "-" & nvl(rsTemp!����))
            End If
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = nvl(rsTemp!����)
        objCtl.Tag = nvl(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int���� As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int���� = rsTemp!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str����)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str���� = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zl_From��Ա��ȡȱʡ����(ByVal lng��ԱID As Long, ByRef str���� As String, ByRef str���� As String, ByRef lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-12-16 10:56:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    gstrSQL = "Select b.Id,b.����,b.���� From ������Ա A,���ű�  b Where a.����id=b.Id And a.ȱʡ=1 And a.��Աid=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡȱʡ������Ϣ", lng��ԱID)
    If rsTemp.EOF Then Exit Function
    lng����ID = nvl(rsTemp!id): str���� = nvl(rsTemp!����): str���� = nvl(rsTemp!����)
    zl_From��Ա��ȡȱʡ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
