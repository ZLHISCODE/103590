Attribute VB_Name = "mdlPatiCard"
Option Explicit
Public grsҽ�ƿ����  As ADODB.Recordset
Public gObjYLCards As clsCards
Public gObjYLCardObjs As clsCardObjects   '��ǰ������Ч��ҽ�ƿ�
Public gfrmCardMgr As Object
Public gblnNotCloseWindows  As Boolean '���رմ���
Public grsSystem As ADODB.Recordset
Public Function zlInitPatiCards(Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-05-23 17:54:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int�Զ���ȡ As Integer, bln���� As Boolean, str���� As String, objCard As clsCard
    Dim int�Զ���ȡ��� As Integer, str�������� As String
    Dim objBrushCards As Object
    
    Err = 0: On Error GoTo Errhand:
    
    
    Set gObjYLCards = New clsCards
    Set gObjYLCardObjs = New clsCardObjects
    
    Set grsҽ�ƿ���� = Nothing: Set grsStatic.rs���ѿ��ӿ� = Nothing
    
    Set rsTemp = zlGetҽ�ƿ����(cnOracle)
    With rsTemp
        '���ƿ�(�����ѿ�)
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & nvl(!����), "�Զ���ȡ", "0"))
            int�Զ���ȡ��� = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & nvl(!����), "�Զ���ȡ���", "300"))
            bln���� = Val(nvl(rsTemp!�Ƿ�����)) = 1
            
            '90875,���ϴ�,2016/1/22:֤�����Ͷ�����
            If bln���� Then
                If Val(nvl(rsTemp!�Ƿ�����)) = 1 Or Val(nvl(rsTemp!�Ƿ�֤��)) = 1 Then   '���ƿ�,������
                    bln���� = True
                Else
                    '�����:54098
                    If (nvl(rsTemp!����) Like "*���֤*" Or nvl(rsTemp!����) Like "*IC��*") And Val(nvl(rsTemp!�Ƿ�̶�)) = 1 And nvl(rsTemp!����) = "" Then
                        bln���� = True
                    Else
                        bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\ҽ�ƿ�\" & nvl(!����), "����", "0")) = 1
                    End If
                End If
            End If
            str���� = Trim(nvl(rsTemp!����))
            'ID,����,����,����,ǰ׺�ı�,���ų���,ȱʡ��־,�Ƿ�̶�,�Ƿ��ϸ����,�Ƿ�ˢ��,�Ƿ�����,�Ƿ�����ʻ�,�Ƿ�ȫ��,����,��ע,�ض���Ŀ,���㷽ʽ,�Ƿ�����
            '77872,���ϴ�,2014/10/28:�Ƿ�֧��ת�ʼ�����
            '85565,���ϴ�,2015/7/10:��������
            '90875,���ϴ�,2016/1/22:�Ƿ�֤��
            '103310,���ϴ�,2016/12/7:���ź����ӻس���λ
            Set objCard = New clsCard
            With objCard
                .������ = EM_CardType_Square
                .�ӿ���� = nvl(rsTemp!id)
                .�ӿڱ��� = nvl(rsTemp!����)
                .���� = nvl(rsTemp!����)
                .���� = nvl(rsTemp!����)
                .ǰ׺�ı� = nvl(rsTemp!ǰ׺�ı�)
                .���ų��� = Val(nvl(rsTemp!���ų���)) + Val(nvl(rsTemp!�豸�Ƿ����ûس�))
                .ȱʡ��־ = Val(nvl(rsTemp!ȱʡ��־)) = 1
                .ϵͳ = Val(nvl(rsTemp!�Ƿ�̶�)) = 1
                .�Ƿ��ϸ���� = Val(nvl(rsTemp!�Ƿ��ϸ����)) = 1
                .�Ƿ��Զ���ȡ = int�Զ���ȡ
                .�Զ���ȡ��� = int�Զ���ȡ���
                .���ƿ� = Val(nvl(rsTemp!�Ƿ�����)) = 1
                .�Ƿ�����ʻ� = Val(nvl(rsTemp!�Ƿ�����ʻ�)) = 1
                .�Ƿ�ȫ�� = Val(nvl(rsTemp!�Ƿ�ȫ��)) = 1
                .�����ظ�ʹ�� = Val(nvl(rsTemp!�Ƿ��ظ�ʹ��)) = 1
                .���㷽ʽ = nvl(rsTemp!���㷽ʽ)
                .�ӿڳ����� = nvl(rsTemp!����)
                .�ض���Ŀ = nvl(rsTemp!�ض���Ŀ)
                .���� = bln����
                .��ע = nvl(rsTemp!��ע)
                .�������Ĺ��� = nvl(rsTemp!��������)
                .�Ƿ����� = Val(nvl(rsTemp!�Ƿ�����)) = 1
                .���볤�� = Val(nvl(rsTemp!���볤��))
                .���볤������ = Val(nvl(rsTemp!���볤������))
                .������� = Val(nvl(rsTemp!�������))
                .������������ = Val(nvl(rsTemp!������������))
                .�Ƿ�ȱʡ���� = Val(nvl(rsTemp!�Ƿ�ȱʡ����)) = 1
                .�Ƿ��ƿ� = Val(nvl(rsTemp!�Ƿ��ƿ�)) = 1   '56615
                .�Ƿ񷢿� = Val(nvl(rsTemp!�Ƿ񷢿�)) = 1 Or .���ƿ�
                .�Ƿ�д�� = Val(nvl(rsTemp!�Ƿ�д��)) = 1
                .�Ƿ�ģ������ = Val(nvl(rsTemp!�Ƿ�ģ������)) = 1
                .�Ƿ�ת�ʼ����� = Val(nvl(rsTemp!�Ƿ�ת�ʼ�����)) = 1
                str�������� = nvl(rsTemp!��������, "1000")
                .�Ƿ�ˢ�� = Mid(str��������, 1, 1) = 1
                .�Ƿ�ɨ�� = Mid(str��������, 2, 1) = 1
                .�Ƿ�Ӵ�ʽ���� = Mid(str��������, 3, 1) = 1
                .�Ƿ�ǽӴ�ʽ���� = Mid(str��������, 4, 1) = 1
                .�Ƿ�֤�� = Val(nvl(rsTemp!�Ƿ�֤��)) = 1
                .�Ƿ�ֿ����� = Val(nvl(rsTemp!�Ƿ�ֿ�����)) = 1
                .���͵��ýӿ� = Val(nvl(rsTemp!���͵��ýӿ�)) = 1
                .�Ƿ��˿��鿨 = Val(nvl(rsTemp!�Ƿ��˿��鿨)) = 1
                .�豸�Ƿ����ûس� = Val(nvl(rsTemp!�豸�Ƿ����ûس�)) = 1
                .�Ƿ�ȱʡ���� = Val(nvl(rsTemp!�Ƿ�ȱʡ����)) = 1
            End With
            gObjYLCards.Add objCard, "K" & objCard.�ӿ����
            If zlCreatePatiCardObject(objCard, objBrushCards) Then
                gObjYLCardObjs.Add objBrushCards, objCard.���ƿ�, objCard.�ӿ����, objCard, False, "K" & objCard.�ӿ����
            End If
            .MoveNext
        Loop
    End With
    
    Set rsTemp = zlGet���ѿ��ӿ�(cnOracle)
    With rsTemp
        '���ƿ�(�����ѿ�)
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & nvl(!���), "�Զ���ȡ", "0"))
            int�Զ���ȡ��� = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & nvl(!���), "�Զ���ȡ���", "300"))
            bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & nvl(!���), "����", "0")) = 1
            
            '���,����,���㷽ʽ,nvl(���ƿ�,0)  as ���ƿ�,ǰ׺�ı�,���ų���,����,ϵͳ,�Ƿ�����
            str���� = Trim(nvl(rsTemp!����))
            Set objCard = New clsCard
            With objCard
                .������ = EM_CardType_Consume
                .�ӿ���� = nvl(rsTemp!���)
                .�ӿڱ��� = nvl(rsTemp!���)
                .���� = Left(nvl(rsTemp!����), 1)   'Ĭ��ȡ��һ��
                .���� = nvl(rsTemp!����)
                .ǰ׺�ı� = nvl(rsTemp!ǰ׺�ı�)
                .���ų��� = Val(nvl(rsTemp!���ų���))
                .ϵͳ = Val(nvl(rsTemp!ϵͳ)) = 1
                .�Ƿ��ϸ���� = False
                .�Ƿ��Զ���ȡ = int�Զ���ȡ
                .�Զ���ȡ��� = int�Զ���ȡ���
                .���ƿ� = Val(nvl(rsTemp!���ƿ�)) = 1
                .�Ƿ�����ʻ� = True 'Not (Val(Nvl(rsTemp!���ƿ�)) = 1)
                .�Ƿ�ȫ�� = Val(nvl(rsTemp!�Ƿ�ȫ��)) = 1
                .���㷽ʽ = nvl(rsTemp!���㷽ʽ)
                .�ӿڳ����� = nvl(rsTemp!����)
                .�ض���Ŀ = ""
                .���� = bln����
                .�����ظ�ʹ�� = True
                .��ע = ""
                .�������Ĺ��� = nvl(rsTemp!�Ƿ�����)
                .���ѿ� = True
                .�Ƿ����� = Val(nvl(rsTemp!�Ƿ�����)) = 1
                .���볤�� = Val(nvl(rsTemp!���볤��))
                .���볤������ = Val(nvl(rsTemp!���볤������))
                .������� = Val(nvl(rsTemp!�������))
                .������������ = Val(nvl(rsTemp!������������))
                .�Ƿ�ȱʡ���� = Val(nvl(rsTemp!�Ƿ�ȱʡ����)) = 1
                .�Ƿ��ƿ� = Val(nvl(rsTemp!�Ƿ��ƿ�)) = 1   '56615
                .�Ƿ񷢿� = Val(nvl(rsTemp!�Ƿ񷢿�)) = 1 Or .���ƿ�
                .�Ƿ�д�� = Val(nvl(rsTemp!�Ƿ�д��)) = 1
                
                str�������� = nvl(rsTemp!��������, "1000")
                .�Ƿ�ˢ�� = Mid(str��������, 1, 1) = 1
                .�Ƿ�ɨ�� = Mid(str��������, 2, 1) = 1
                .�Ƿ�Ӵ�ʽ���� = Mid(str��������, 3, 1) = 1
                .�Ƿ�ǽӴ�ʽ���� = Mid(str��������, 4, 1) = 1
            End With
            gObjYLCards.Add objCard, "X" & objCard.�ӿ����
            If zlCreatePatiCardObject(objCard, objBrushCards) Then
                gObjYLCardObjs.Add objBrushCards, objCard.���ƿ�, objCard.�ӿ����, objCard, True, "X" & objCard.�ӿ����
            End If
            .MoveNext
        Loop
    End With
    zlInitPatiCards = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlGetYLCardObjs(ByRef objYlCardObjects As clsCardObjects) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����
    '����:objYlCardObjects-���ؿ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-23 13:59:24
    '˵��:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gObjYLCardObjs Is Nothing Then
        Set objYlCardObjects = gObjYLCardObjs
        zlGetYLCardObjs = True
        Exit Function
    End If
    If gcnOracle.State <> 1 Then Exit Function
    If zlInitPatiCards = False Then Exit Function
    Set objYlCardObjects = gObjYLCardObjs
    zlGetYLCardObjs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetCards_YL(ByRef objCards As clsCards) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����
    '����:objCards-ҽ�ƿ�������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-04-23 12:03:26
    '˵��:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gObjYLCards Is Nothing Then
        Set objCards = gObjYLCards: zlGetCards_YL = True: Exit Function
    End If
    If zlInitPatiCards = False Then Exit Function
    Set objCards = gObjYLCards
    zlGetCards_YL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSystemNo(ByVal lngSys As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����ϵͳ��(��:���2100,�����HIS��Ϊ100��101)
    '����:���ع���ĺ���(û�й����,ֱ�ӷ��ص�ǰϵͳ��)
    '����:���˺�
    '����:2011-09-21 16:13:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Int(lngSys / 100) = 1 Then zlGetSystemNo = lngSys: Exit Function
    
    On Error GoTo errHandle
    
   gstrSQL = "Select ���,����,�����,������  From zltools.zlsystems "
    If grsSystem Is Nothing Then
        Set grsSystem = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ϵͳ��")
    ElseIf grsSystem.State <> 1 Then
        Set grsSystem = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ϵͳ��")
    End If
    grsSystem.Filter = "���=" & lngSys
    If grsSystem.EOF = False Then
        If Val(nvl(grsSystem!�����)) <> 0 Then zlGetSystemNo = grsSystem!�����
    End If
    grsSystem.Filter = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetPatiDayMoney(lng����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵��췢���ķ����ܶ�
    '����:��ȡ���˵��췢���ķ����ܶ�
    '����:���˺�
    '����:2011-06-23 10:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng����ID As Long) As Double
'����:��ȡָ�����˵Ļ��۵����ϼ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '���ʱ�����������סԺ���۷���
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara("���ʱ�����������סԺ���۷���", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(��ҳID,0) = (Select Nvl(��ҳID,0) From ������Ϣ Where ����ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־=2" & strWhere
    Else
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� " & _
        "   From ������ü�¼  " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]  and �����־<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(���۷��úϼ�,0)) as ���۷��úϼ�  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ�����˵Ļ����ܶ�", lng����ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetҽ�ƿ����(Optional cnOracle As ADODB.Connection) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����
    '����:����ҽ�ƿ����ļ�¼��
    '����:���˺�
    '����:2011-05-23 17:25:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDatabase As Object, objTemp As clsDataBase
    
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    '�����:51072,56615:�Ƿ��ƿ�,�Ƿ񷢿�,�Ƿ�д��
    '�Ȼ��浽����
    '77872,���ϴ�,2014/10/28:�Ƿ�֧��ת�ʼ�����
    '90875,���ϴ�,2016/1/22:�Ƿ�֤������
    '104238:���ϴ���2017/2/15��ҽ�ƿ�������ӷ������ſ���
    gstrSQL = "" & _
    "   Select A.Id, A.����, A.����, A.����, A.ǰ׺�ı�, A.���ų���, A.ȱʡ��־, A.�Ƿ�̶�, A.�Ƿ��ϸ����, " & _
    "           nvl(A.�Ƿ�����,0) as �Ƿ�����, nvl(A.�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(A.�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(A.�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� , nvl(A.��������,0) as ��������, " & _
    "           nvl(A.���볤��,10) as ���볤��,nvl(A.���볤������,0) as ���볤������,nvl(A.�������,0) as �������," & _
    "           nvl(A.�Ƿ�����,0) as �Ƿ�����,A.����, A.��ע, A.�ض���Ŀ, A.���㷽ʽ, A.�Ƿ�����, A.��������,Nvl(A.������������,0) as ������������,Nvl(A.�Ƿ�ȱʡ����,0) as �Ƿ�ȱʡ����," & _
    "           nvl(A.�Ƿ�ģ������,0) as �Ƿ�ģ������,nvl(A.�Ƿ��ƿ�,0) as �Ƿ��ƿ�, decode(nvl(A.�Ƿ�����,0),1,1,nvl(A.�Ƿ񷢿�,0)) as �Ƿ񷢿�, nvl(A.�Ƿ�д��,0) as �Ƿ�д��," & _
    "           B.����  as ��������, nvl(A.�Ƿ�����,0) as �Ƿ�����,nvl(�Ƿ�֤��,0) as �Ƿ�֤��, " & _
    "           nvl(A.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����, nvl(A.��������,'1000') as ��������, " & _
    "           nvl(A.�Ƿ�ֿ�����,0) as �Ƿ�ֿ�����,nvl(A.���͵��ýӿ�,0) as ���͵��ýӿ�," & _
    "           Nvl(a.�Ƿ��˿��鿨,0) As �Ƿ��˿��鿨, A.�豸�Ƿ����ûس�, nvl(A.��������,0) as ��������," & _
    "           Nvl(A.�Ƿ�ȱʡ����,0) as �Ƿ�ȱʡ���� " & _
    "    From ҽ�ƿ���� A,���㷽ʽ B" & _
    "    Where A.���㷽ʽ=B.����(+)" & _
    "    Order by ����"
    
    If grsҽ�ƿ���� Is Nothing Then
        Set grsҽ�ƿ���� = objDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ѿ��ӿ� ")
    ElseIf grsҽ�ƿ����.State <> 1 Then
        Set grsҽ�ƿ���� = objDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ѿ��ӿ� ")
    End If
    grsҽ�ƿ����.Filter = 0
    Set zlGetҽ�ƿ���� = grsҽ�ƿ����
    Set objDatabase = Nothing
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        Set objTemp = Nothing: Set objDatabase = Nothing
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objTemp = Nothing: Set objDatabase = Nothing
End Function

Public Function zlCreatePatiCardObject(ByVal objCard As clsCard, ByRef objCardObject As Object, Optional blnAdviceSend As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����Ķ���
    '����:objCardObject-�������Ķ���
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-25 10:47:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String, strHead As String
    If Not objCard.���� And Not blnAdviceSend Then
        Set objCardObject = Nothing: Exit Function
    End If
    
    '����豸�Ƿ�����
    strHead = IIf(objCard.���ѿ�, "", "zl9Card_")
    If objCard.�ӿڳ����� = "" Then
        '99858:���ϴ�,2016/9/2,�����˻�ҽ�ƿ������������ѿ������нӿڲ���
        If objCard.���ѿ� And objCard.���ƿ� Then
            Set objCardObject = New clsSimulateSquareCard: zlCreatePatiCardObject = True: Exit Function
        ElseIf Not objCard.�Ƿ�����ʻ� Then
            Set objCardObject = New clsOwnerCardObject: zlCreatePatiCardObject = True: Exit Function
        End If
        MsgBox objCard.���� & "δ���ýӿڲ���������" & IIf(objCard.���ѿ�, "�����ѿ�����", "��ҽ�ƿ�������") & "�����ò�����!"
        Exit Function
    End If
    strCommpentName = GetCardComponentsStr(objCard.�ӿڳ�����, strHead)
    Err = 0: On Error Resume Next
    Set objCardObject = CreateObject(strCommpentName)
    If Err <> 0 Then
        ShowMsgbox "����:" & objCard.�ӿڱ��� & "-" & objCard.���� & "( " & strCommpentName & ")����ʧ��,����ϵͳ����Ա��ϵ!" & vbCrLf & "��ϸ����ϢΪ:" & Err.Description
        Call WritLog("mdlCardSquare.zlCreatePatiCardObject", "", "����:" & objCard.�ӿڱ��� & "-" & objCard.���� & "����ʧ��!��ϸ����ϢΪ:" & Err.Description)
        Exit Function
    End If
    zlCreatePatiCardObject = True
End Function

Public Function zlGetComponentObject(ByVal lng�����ID As Long, _
     Optional bln���ѿ� As Boolean = False) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ŀ�����
    '���:lng�����ID-�����ID
    '        bln���ѿ�
    '����:
    '����:
    '����:���˺�
    '����:2011-06-25 23:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objYlCardObjs As clsCardObjects
    strKey = IIf(bln���ѿ�, "X", "K") & lng�����ID
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Err = 0: On Error Resume Next
    Set zlGetComponentObject = objYlCardObjs(strKey).CardObject
    If Err <> 0 Then
        Err.Clear: On Error GoTo 0
    End If
End Function
Public Function zlGetClsCardObject(ByVal lng�����ID As Long, _
     Optional bln���ѿ� As Boolean = False) As clsCardObject
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ŀ�����
    '���:lng�����ID-�����ID
    '        bln���ѿ�
    '       blnChkExeObject-���ִ�ж���
    '       blnChkInitComents-����ʼ����
    '����:
    '����:�Ϸ�,����clsCardObject����
    '����:���˺�
    '����:2011-06-25 23:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    strKey = IIf(bln���ѿ�, "X", "K") & lng�����ID
    Err = 0: On Error Resume Next
    Set zlGetClsCardObject = objYlCardObjs(strKey)
    If Err <> 0 Then
        Err.Clear: On Error GoTo 0
    End If
End Function
   
Public Function zlItemsMoney(ByVal strIDs As String, Optional ByVal strPriceGrade As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����Ŀ������
    '���:����ID�����ö��ŷ���
    '����:��ؼ۸�
    '����:���˺�
    '����:2011-05-31 15:24:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, j As Long
    Dim strSubTable As String, varData() As Variant
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    Call zlGetSubTable(0, strIDs, strSubTable, varData)
    '�۸�ȼ�
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [" & UBound(varData) + 2 & "]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [" & UBound(varData) + 2 & "]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    
    strSubTable = " With �Һ���Ŀ as ( " & strSubTable & ") "
    
    strSQL = strSubTable & _
        "   Select  /*+ rule */  1 as ����,A.���,A.ID as ����ID,0 as ����ĿID, " & _
        "               A.���� as ��Ŀ����,A.���� as ��Ŀ����, A.���㵥λ,A.���ηѱ�," & _
        "               1 as ����,B.�ּ� as ����, " & _
        "               C.ID as ������ĿID,C.���� as ������Ŀ, " & _
        "               C.���� as �������,C.�վݷ�Ŀ" & _
        "   From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�Һ���Ŀ M" & _
        "   Where A.ID =M.ID and B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID  " & _
        "               And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
                strWherePriceGrade & vbNewLine & _
        "   Union ALL " & _
        "   Select 2 as ����,A.���,D.����ID,A.ID as ��ĿID, " & _
        "               A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
        "               D.�������� as ����,B.�ּ� as ����, " & _
        "               C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ" & _
        "   From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D,�Һ���Ŀ M" & _
        "   Where A.ID=D.����ID And D.����ID =M.ID And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID   " & _
        "           And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
                strWherePriceGrade
    strSQL = "Select /*+ RULE */  * From (" & strSQL & ")        "
    
    ReDim Preserve varData(UBound(varData) + 1)
    varData(UBound(varData)) = strPriceGrade
    Set zlItemsMoney = zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡ�Һż۸�", varData)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetRecodersFieldIns(ByVal rsTemp As ADODB.Recordset, ByVal strFieldNames As String, ByRef cllData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����¼���е��ֶμ�
    '���:rsTemp-ָ���ļ�¼��
    '        strFields-ָ�����ֶ�,����:��ĿID,����ID
    '����:cllData-����ָ���ļ�(��:��ĿID,����IDΪ��������)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-31 17:38:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varFields As Variant, i As Long
    Dim strFields() As String, strTemp As String
    varFields = Split(strFieldNames, ",")
    ReDim Preserve strFields(0 To UBound(varFields)) As String
    With rsTemp
        Do While Not .EOF
            For i = 0 To UBound(varFields)
                If InStr(1, strFields(i) & ",", "," & .Fields(varFields(i)).value & ",") = 0 Then
                    strFields(i) = strFields(i) & "," & .Fields(varFields(i)).value
                End If
            Next
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        Set cllData = New Collection
        For i = 0 To UBound(varFields)
            strTemp = strFields(i)
            If Trim(strTemp) <> "" Then strTemp = Mid(strTemp, 2)
            cllData.Add strTemp, varFields(i)
        Next
    End With
End Function
Public Function zlGetSubTable(ByVal bytType As Byte, ByVal strValues_IN As String, _
    strSubTable As String, varPara() As Variant, Optional intParaStep As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ַ����ֽ���ӱ��ѯ(��Num2list),������20����
    '���:bytType: 0-Num2List;1-Str2List;2-Num2List2;3Str2List2
    '       strValues_IN:bytType=0,1ʱ,֮���ö��ŷ���
    '                            bytType=2,3ʱ,��֮����:������֮����,����:��:����:22,����:22
    '       varParaStep_in:��������ʼ����
    '����:varPara-������0-20������
    '����:�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-06-01 10:50:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, j As Long
    Dim strTemp As String, strSplit As String
    Dim varValue As Variant
    Dim strTable As String
    i = intParaStep: strSplit = ","
    If bytType = 0 Or bytType = 2 Then
        Do While strValues_IN <> ""
            ReDim Preserve varPara(0 To i) As Variant
            If Len(strValues_IN) > 4000 Then
                j = InStr(IIf(bytType = 0, 3982, 3958), strValues_IN, strSplit)
                strTemp = Mid(strValues_IN, 1, j - 1): strValues_IN = Mid(strValues_IN, j + 1)
                varPara(i) = strTemp
            Else
                strTemp = strValues_IN
                varPara(i) = strTemp
                strValues_IN = ""
            End If
            i = i + 1
        Loop
    Else
        varValue = Split(strValues_IN, strSplit)
        strTemp = ""
        For j = 0 To UBound(varValue)
              If zlCommFun.ActualLen(strTemp & "," & varValue(j)) > 4000 Then
                  ReDim Preserve varPara(0 To i) As Variant
                  varPara(i) = Mid(strTemp, 2): i = i + 1
                  strTemp = ""
              End If
              strTemp = strTemp & "," & varValue(j)
        Next
        If strTemp <> "" Then
            ReDim Preserve varPara(0 To i) As Variant
            varPara(i) = Mid(strTemp, 2)
        End If
    End If
    For i = intParaStep To UBound(varPara)
        If varPara(i) <> "" Then
            j = i + 1
            If bytType = 0 Then
                strTable = strTable & " Union All Select Column_Value as ID From Table( f_Num2list([" & j & "])) "
            ElseIf bytType = 1 Then
                strTable = strTable & " Union All Select Column_Value From Table( f_str2list([" & j & "])) "
            ElseIf bytType = 2 Then
                strTable = strTable & " Union All Select  C1,C2 From Table( f_Num2list2([" & j & "])) "
            Else
                strTable = strTable & " Union All Select  C1,C2 From Table( f_Str2list2([" & j & "])) "
            End If
        End If
    Next
    If strTable = "" Then Exit Function
    strSubTable = "Select distinct  * From ( " & Mid(strTable, 11) & ")"
    zlGetSubTable = True
End Function

Public Function zlGetActualMoney(ByVal str�ѱ� As String, ByVal lng����ID As Long, ByVal dblӦ�� As Double, ByVal lng�շ�ϸĿID As Long) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ķѱ��������Ŀ���շ���Ŀ,����ָ������ʵ���տ���
    '���:str�ѱ�-�ѱ�
    '        lng����ID-������ĿID
    '        dblӦ��-Ӧ�ս��ֵ
    '����:
    '����:ʵ��Ӧ�յĽ��
    '����:���˺�
    '����:2011-06-02 11:50:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4])  as ʵ�ս�� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng����ID, dblӦ��)
    If rsTmp.EOF Then
        zlGetActualMoney = dblӦ��
    Else
        zlGetActualMoney = Round(Val(Split(nvl(rsTmp!ʵ�ս��) & ":", ":")(1)), 5)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetIDKindStr(Optional strIDKindStr As String = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0", Optional blnOnlyAccouct As Boolean = False, Optional objSquare As clsCardSquare) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Чҽ�ƿ��ַ���
    '���:strIDKindStr    String  IN
    '�����ָ�ʽ:
    'һ����ȱʡ��:����1|ȫ��1|������־1;��. ;����n|ȫ��n|������־n
    '��һ���Ǹýӿڷ��صĸ�ʽ:����|ȫ��|������־|�����ID|���ų���|ȱʡ��־|�Ƿ�����ʻ�;��
    '����:
    '����: ����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)|�Ƿ�ɨ��|�Ƿ�Ӵ�ʽ����|�Ƿ�ǽӴ�ʽ����;��
    '        ����:�����ID|�����Ǳ������ӵ�,�ɵ����߸��������ȷ��.
    '       ����:��|���֤��|0|0|18|0;IC|IC����|1|0|8|0;��|�����|0|0|0|0;��|���￨|0|0|0|1;��|���п�|0|0|10|0
    '      ���ִ���ʱ,���ؿ�
    '����:���˺�
    '����:2011-06-14 14:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim strNewIdKindStr As String, strTemp As String
    Dim lngMaxLen As Long, blnPassText As Boolean
    Dim blnExists As Boolean '�Ƿ����ģ������
    '77076,Ƚ����,2014-8-25,ͬʱ��ҽ�ƿ����Ź���Ͳ�����Ϣ����������Ժ����,�ڲ�����Ϣ�����д򿪵Ǽ�,ҽ�ƿ����Ź������Զ��ر�
    Dim objCard As New Card, objCards As New Cards, objCardSquare As clsCardSquare
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandle
        
    If objSquare Is Nothing Then
        Set objCardSquare = New clsCardSquare
    Else
        Set objCardSquare = objSquare
    End If
    strNewIdKindStr = ""
    varData = Split(strIDKindStr, ";")
    blnExists = False
    '76187,Ƚ����,2014-8-4
    objCardSquare.mblnYLMgr = True
    Set objCards = objCardSquare.zlGetCards(1)
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|||||||", "|")
            blnFind = False
            For Each objCard In objCards
                If objCard.���� = varTemp(1) Then blnFind = True: Exit For
            Next
            If blnFind Then
                '85565,���ϴ�,2015/7/10:��������
                strNewIdKindStr = strNewIdKindStr & ";" & objCard.���� & "|" & objCard.���� & "|" & IIf(objCard.�Ƿ�ˢ��, 0, 1) & _
                                "|" & objCard.�ӿ���� & "|" & objCard.���ų��� & "|" & IIf(objCard.ȱʡ��־, 1, 0) & _
                                "|" & IIf(objCard.�Ƿ�����ʻ�, 1, 0) & "|" & objCard.�������Ĺ��� & _
                                "|" & IIf(objCard.�Ƿ�ɨ��, 1, 0) & "|" & IIf(objCard.�Ƿ�Ӵ�ʽ����, 1, 0) & "|" & IIf(objCard.�Ƿ�ǽӴ�ʽ����, 1, 0)
                strTemp = strTemp & "," & objCard.�ӿ����
            Else
                '����|ȫ��|������־|�����ID(-1����ģ������)|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)|�Ƿ�ɨ��|�Ƿ�Ӵ�ʽ����|�Ƿ�ǽӴ�ʽ����;��
                strNewIdKindStr = strNewIdKindStr & ";" & varTemp(0) & "|" & varTemp(1) & "|" & Val(varTemp(2)) & "|" & varTemp(3) & "|" & varTemp(4) & "|||"
            End If
            If varTemp(1) = "ģ������" And Val(varTemp(3)) < 0 Then blnExists = True
        End If
    Next

    For Each objCard In objCards
        If InStr(1, strTemp & ",", "," & objCard.�ӿ���� & ",") = 0 Then
            If Not blnOnlyAccouct Or (blnOnlyAccouct And objCard.�Ƿ�����ʻ�) Then
                '85565,���ϴ�,2015/7/10:��������
                strNewIdKindStr = strNewIdKindStr & ";" & objCard.���� & "|" & objCard.���� & "|" & IIf(objCard.�Ƿ�ˢ��, 0, 1) & _
                                "|" & objCard.�ӿ���� & "|" & objCard.���ų��� & "|" & IIf(objCard.ȱʡ��־, 1, 0) & _
                                "|" & IIf(objCard.�Ƿ�����ʻ�, 1, 0) & "|" & objCard.�������Ĺ��� & _
                                "|" & IIf(objCard.�Ƿ�ɨ��, 1, 0) & "|" & IIf(objCard.�Ƿ�Ӵ�ʽ����, 1, 0) & "|" & IIf(objCard.�Ƿ�ǽӴ�ʽ����, 1, 0)
            End If
        End If
    Next
    
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    GetIDKindStr = strNewIdKindStr
    Exit Function
errHandle:
    GetIDKindStr = ""
End Function

Private Function IsCreateObject(ByVal str���� As String, Optional strHead As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ񴴽��ɹ�
    '���:strHead-���������ļ�ͷ:����:ҽ�ƿ��Բ�����:zl9Card_��ͷ.
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 15:43:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Object
    If str���� = "" Then IsCreateObject = True: Exit Function
    str���� = GetCardComponentsStr(str����)
    Err = 0: On Error Resume Next
    Set objTemp = CreateObject(str����)
    If Err <> 0 Then
       Err.Clear: On Error GoTo 0
       Set objTemp = Nothing
       IsCreateObject = False: Exit Function
    End If
    Set objTemp = Nothing
    IsCreateObject = True
End Function
Private Function GetCardComponentsStr(ByVal str���� As String, Optional strHead As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:
    '����:���˺�
    '����:2011-06-22 13:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If str���� = "" Then GetCardComponentsStr = "": Exit Function
    If strHead <> "" Then
        '�п�ͷ����,��Ҫ��鴫���Ƿ�����ⲿ��
        If str���� Like strHead & "*" Then
            str���� = str���� & "." & "cls" & Mid(str����, Len(strHead) + 1)
        Else
            str���� = strHead & str���� & "." & "cls" & Replace(Replace(UCase(str����), "ZL9", ""), "ZL", "")
        End If
    Else
        str���� = str���� & "." & "cls" & Replace(Replace(UCase(str����), "ZL9", ""), "ZL", "")
    End If
    GetCardComponentsStr = str����
End Function
Public Function zlGetPrivFuns(ByVal lngModule As Long, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ģ���Ȩ��
    '����:����Ȩ�޴�
    '����:���˺�
    '����:2015-06-03 09:46:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String, objDatabase As clsDataBase
    On Error GoTo errHandle
    If cnOracle Is Nothing Then
        strPrivs = ";" & GetPrivFunc(glngSys, lngModule) & ";"
    Else
        Set objDatabase = New clsDataBase
        Call objDatabase.InitCommon(cnOracle)
        strPrivs = ";" & objDatabase.GetPrivFunc(glngSys, lngModule) & ";"
        Set objDatabase = Nothing
    End If
    zlGetPrivFuns = strPrivs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetAvailabilityCardType(Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ����վ��Ч��֧����,�������п�;���ѿ���
    '����:��ʽ:��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�|�Ƿ�����|�Ƿ�ȫ��|�Ƿ�ɨ��|�Ƿ�Ӵ�ʽ����|�Ƿ�ǽӴ�ʽ����;��
    '����:���˺�
    '����:2011-06-14 15:16:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int�Զ���ȡ As Integer, strҽ�ƿ� As String
    Dim bln���� As Boolean, strCardStr As String
    Dim objCard As clsCard, i As Long, blnAdd As Boolean
    Dim strPrivs As String, bln֧�������ӿ� As Boolean  '-False,��ʾֻ֧�����ѿ�;True:֧�����ѿ�;���п���
    Dim objYlCardObjs As clsCardObjects
    Dim objDatabase As clsDataBase
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    On Error GoTo errHandle
    
    '�Ƿ���һ��ͨ���Ѳ���������ģ��
    strPrivs = zlGetPrivFuns(1151, cnOracle)
    
    bln֧�������ӿ� = InStr(1, strPrivs, ";�����ӿ�����;") > 0
    
    For i = 1 To objYlCardObjs.count
        If objYlCardObjs(i).CardPreporty.���� And objYlCardObjs(i).CardPreporty.�Ƿ�����ʻ� Then
            If Not objYlCardObjs(i).CardObject Is Nothing Then
                blnAdd = True
                If bln֧�������ӿ� = False Then
                    blnAdd = False
                    If objYlCardObjs(i).CardPreporty.���ѿ� And _
                        objYlCardObjs(i).CardPreporty.���ƿ� Then
                        blnAdd = True
                    End If
                End If
                If blnAdd Then
                    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�|�Ƿ�����|�Ƿ�ȫ��|�Ƿ�ɨ��|�Ƿ�Ӵ�ʽ����|�Ƿ�ǽӴ�ʽ����;��
                    '85565,���ϴ�,2015/7/10:��������
                    Set objCard = objYlCardObjs(i).CardPreporty
                    strCardStr = strCardStr & ";" & objCard.���� & "|" & objCard.���� & "|" & IIf(objCard.�Ƿ�ˢ��, 0, 1)
                    strCardStr = strCardStr & "|" & objCard.�ӿ���� & "|" & objCard.���ų���
                    strCardStr = strCardStr & "|" & IIf(objCard.���ѿ�, 1, 0) & "|" & objCard.���㷽ʽ
                    strCardStr = strCardStr & "|" & objCard.�������Ĺ��� & "|" & IIf(objCard.���ƿ�, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.�Ƿ�����, 1, 0) & "|" & IIf(objCard.�Ƿ�ȫ��, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.�Ƿ�ɨ��, 1, 0) & "|" & IIf(objCard.�Ƿ�Ӵ�ʽ����, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.�Ƿ�ǽӴ�ʽ����, 1, 0) & "|" & IIf(objCard.�Ƿ��˿��鿨, 1, 0)
                    '����:50120
                End If
            End If
        End If
    Next
    If strCardStr <> "" Then strCardStr = Mid(strCardStr, 2)
     GetAvailabilityCardType = strCardStr
    Exit Function
errHandle:
     GetAvailabilityCardType = ""
End Function

Public Function GetIDKindCardTypeID(ByVal strIDKindStr As String, ByVal strIDKind As String, _
    ByRef lngCardTypeID As Long, ByRef lngCardLen As Long, _
    Optional ByRef strKindName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ҫ�Ǹ�����IDKind�����е�IDKind����ȡ��Ӧ�����ID,�Ա������صĲ�����Ϣ:
    '���:strIDKindStr  -ȱʡ��StrIDKindStr:��ʽ:����|ȫ��|������־|�����ID|���ų���;����
    '       ��|���֤��|0|0|18;IC|IC����|1|0|8;��|�����|0|0|0;��|���￨|0|0|0;��|���п�|0|0|10
    '       strIDKind-����Ϊȱʡ������;Ҳ����Ϊ����(������0��N):������,�ظ�ʱ,ָ���һ��.
    '����:lngCardTypeID- �����ID
    '       lngCardLen- ���ų���
    '       strKindName-����
    '����:Boolean ����    �ɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-06-14 16:11:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, blnIndex As Boolean, i As Long
    
    On Error GoTo errHandle
    varData = Split(strIDKindStr, ";")
    lngCardTypeID = -1: lngCardLen = -1
    blnIndex = IsNumeric(strIDKind)
    For i = 0 To UBound(varData)
            If blnIndex Then
                If i = Val(strIDKind) Then
                    '��ʽ:����|ȫ��|������־|�����ID|���ų���;��
                    varTemp = Split(varData(i) & "|||||", "|")
                    lngCardTypeID = Val(varTemp(3))
                    lngCardLen = Val(varTemp(4))
                    strKindName = Trim(varTemp(1))
                    Exit For
                End If
            Else
                    varTemp = Split(varData(i) & "|||||", "|")
                    If varTemp(1) = strIDKind Then
                        lngCardTypeID = Val(varTemp(3))
                        lngCardLen = Val(varTemp(4))
                        strKindName = Trim(varTemp(1))
                        Exit For
                    End If
            End If
    Next
    If lngCardTypeID >= 0 Then
         GetIDKindCardTypeID = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCardFindPati( _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ��ţ�ģ�����Ҳ���
    '        strCardNo-����
    '        blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '����:strErrMsg-���صĴ�����Ϣ
    '        lng����ID-���صĲ���ID
    '        strCardPass-���ؿ��ŵ�����
    '        lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str�����ID As String, str����ID As String, lngTemp As Long
    Dim objDatabase As Object, objTemp As clsDataBase
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    
    'ģ������
    '76020,Ƚ����,2014-7-30,��������ȡ���˿������֧��ģ�����ҡ���ͣ�õĳֿ�����Ϣ
    '114161:���ϴ�,2017/11/7,��ʧ��Ч����ΪNull��0ʱ����ʾ��ʧһֱ��Ч
    strSQL = "" & _
            " Select a.�����id, a.����id, a.����," & _
            "       Case" & _
            "         When Nvl(a.״̬, 0) = 1" & _
            "           And (Nvl(b.��Ч����, 0)  = 0 Or Nvl(a.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.��Ч����, 0) > Sysdate) Then 1" & _
            "         When Nvl(a.״̬, 0) = 2 Then 2" & _
            "         Else 0" & _
            "       End As ״̬" & _
            " From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ B, ҽ�ƿ���� C" & _
            " Where a.�����id = c.Id And Nvl(c.�Ƿ�ģ������, 0) = 1" & _
            "      And a.���� = [2] And a.��ʧ��ʽ = b.����(+)" & _
            "      And Nvl(c.�Ƿ�����, 0) = 1" & _
            " Order By ״̬"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", lngCardTypeID, strCardNo)
    
    Set objDatabase = Nothing: Set objTemp = Nothing
    If rsTemp.EOF Then Exit Function
    rsTemp.Filter = "״̬=0"
    '0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
    If rsTemp.RecordCount = 1 Then
        '��һ����ֱ�ӷ���
        lng����ID = Val(nvl(rsTemp!����ID))
        strCardPassWord = nvl(rsTemp!����)
        rsTemp.Close: Set rsTemp = Nothing
        GetCardFindPati = True: Exit Function
    End If
    If rsTemp.RecordCount = 0 Then Exit Function
    '����
    rsTemp.MoveFirst
    With rsTemp
        str����ID = ""
        Do While Not .EOF
            lngTemp = Val(nvl(!�����id))
            If lngTemp <> 0 Then
                If InStr(1, str�����ID & ",", "," & lngTemp & ",") = 0 Then str�����ID = str�����ID & "," & lngTemp
                If InStr(1, str����ID & ",", "," & Val(nvl(!����ID)) & ",") = 0 Then str����ID = str����ID & "," & Val(nvl(!����ID))
            End If
            .MoveNext
        Loop
        If str����ID <> "" Then str����ID = Mid(str����ID, 2)
        If str�����ID <> "" Then str�����ID = Mid(str�����ID, 2)
        If InStr(1, str����ID, ",") = 0 Then
            .MoveFirst
            lng����ID = Val(nvl(rsTemp!����ID)): lngCardTypeID = Val(nvl(rsTemp!�����id))
            strCardPassWord = nvl(rsTemp!����)
            rsTemp.Close: Set rsTemp = Nothing
            GetCardFindPati = True: Exit Function
        End If
        If frmSelectType.zlSelect(Nothing, str�����ID, lngCardTypeID) = False Then lngCardTypeID = 0: Exit Function
        rsTemp.Filter = "�����ID=" & lngCardTypeID & " And ״̬=0"
        If rsTemp.EOF Then lngCardTypeID = 0: Exit Function
        lng����ID = Val(nvl(rsTemp!����ID))
        strCardPassWord = nvl(rsTemp!����)
        rsTemp.Close: Set rsTemp = Nothing: GetCardFindPati = True: Exit Function
    End With
    '�϶����󣬰���һ����ʾ
     rsTemp.Filter = 0: rsTemp.MoveFirst
     If Val(nvl(rsTemp!״̬)) = 1 Then
        strErrMsg = "����Ϊ" & strCardNo & "�Ѿ�����ʧ!"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(nvl(rsTemp!״̬)) = 2 Then
        strErrMsg = "����Ϊ" & strCardNo & "�Ѿ���ͣ��!"
         If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiID(ByVal strCardType As String, _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional objCtl As Object = Nothing, _
    Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, _
    Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional cnOracle As ADODB.Connection, _
    Optional ByRef blnCertificate As Boolean = False, _
    Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����ҽ�����Ϳ���,��ȡ��Ӧ�Ĳ���ID
    '���:strCardType-�����,���Ϊ����,��Ϊ�����ID,���Ϊ�ַ�,��Ϊ�������
    '       strCardNo-����
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '       frmMain-���õ�������
    '       objCtl-���õĿؼ�
    '       blnShowMergePati-�����ֶ�����������Ĳ���ʱ,�Ƿ���ʾ�ϲ����ܰ�ť
    '       blnOnlyContractPati-ǩԼ����
    '       blnUserCancel-ѡ�����У��û�ѡ����ȡ��
    '       lngShowCardNoTypeID-���˳���������Ϣʱ������ѡ��������ʾ�Ŀ��ŵĿ����ID,0-��ʾ����ʾ���ţ�>0��ʾ��ʾָ����������ID
    '����:strErrMsg-���صĴ�����Ϣ
    '       lng����ID-���صĲ���ID
    '       strCardPass-���ؿ��ŵ�����
    '       lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:��ȡ����ID�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 17:07:51
    '˵��:ֻ�д���ҽ�����Ĳŵ��ô˺���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str�����ID As String, str����ID As Long, lngTemp As Long
    Dim strWhere As String, blnCard As Boolean '
    Dim str������� As String, str��ʶ�� As String
    Dim objDatabase  As Object, objTemp As clsDataBase
    
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    strCardPassWord = "": strErrMsg = ""
    lng����ID = 0
    If strCardType = "" Then Exit Function
    If Val(strCardType) = -1 Then
       GetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID, cnOracle)
       Exit Function
    End If
    
    str������� = ""
    If strCardType Like "*���֤*" Or strCardType Like "*IC��*" Then
        Set rsTemp = zlGetҽ�ƿ����
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Val(nvl(rsTemp!�Ƿ�̶�)) = 1 And Val(nvl(rsTemp!�Ƿ�����)) = 1 Then
                If strCardType Like "*���֤*" And nvl(rsTemp!����) Like "*���֤*" And strCardType <> "��ϵ�����֤" Then
                    str������� = nvl(rsTemp!����)
                     strCardType = Val(nvl(rsTemp!id)): Exit Do
                ElseIf strCardType Like "*IC��*" And nvl(rsTemp!����) Like "*IC��*" Then
                     str������� = nvl(rsTemp!����)
                     strCardType = Val(nvl(rsTemp!id))
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If IsNumeric(strCardType) Then  '�Կ����IDΪ׼
        strSQL = "" & _
        "   Select  A.�����ID, A.����ID, ����, A.״̬, " & _
        "               nvl(��ʧʱ��,to_date('3000-01-01','yyyy-mm-dd'))+nvl(B.��Ч����,0) as ��ʧʱ��," & _
        "               sysdate as ��ǰʱ��  " & _
        "   From ����ҽ�ƿ���Ϣ A,ҽ�ƿ���ʧ��ʽ B" & _
        "   Where  A.�����ID=[1] and A.����=[2] And A.��ʧ��ʽ=B.����(+)"
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", Val(strCardType), strCardNo)
        If Not rsTemp.EOF Then
            lng����ID = Val(nvl(rsTemp!����ID))
            strCardPassWord = nvl(rsTemp!����)
            If Val(nvl(rsTemp!״̬)) = 1 Then
                If Format(rsTemp!��ʧʱ��, "yyyy-mm-dd") <= Format(rsTemp!��ǰʱ��, "yyyy-mm-dd") Then
                    strErrMsg = "����Ϊ" & strCardNo & "�Ѿ�����ʧ!"
                    If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    rsTemp.Close: Set rsTemp = Nothing
                    Exit Function
                End If
            End If
            If Val(nvl(rsTemp!״̬)) = 2 Then
                strErrMsg = "����Ϊ" & strCardNo & "�Ѿ���ͣ��!"
                 If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 rsTemp.Close: Set rsTemp = Nothing
                Exit Function
            End If
            GetPatiID = True
            Exit Function
        End If
        If blnOnlyContractPati Then Exit Function
        
        If str������� = "" Then
              Set rsTemp = zlGetҽ�ƿ����(cnOracle)
              rsTemp.Filter = "ID=" & Val(strCardType) & " And �Ƿ�����=1"
              If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
              If Not rsTemp.EOF Then
                  If Val(nvl(rsTemp!�Ƿ�̶�)) = 1 Then
                      str������� = rsTemp!����
                  End If
              End If
              rsTemp.Filter = 0
          End If
          
        If str������� Like "*���֤*" Then
            strCardType = "���֤"
        ElseIf UCase(str�������) Like "*IC��*" Then
            strCardType = "IC��"
        Else
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
    End If
    
    '90875:���ϴ�,2015/12/16,ҽ�ƿ�֤������
    If blnCertificate Then
        strSQL = "" & _
        "   Select  A.�����ID, A.����ID, ����, A.״̬ " & _
        "   From ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B" & _
        "   Where A.�����ID=B.ID And A.״̬=0 And B.�Ƿ�����=1 And B.����=[1] and A.����=[2] And Nvl(B.�Ƿ�֤��,0)=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", strCardType, strCardNo)
        If rsTemp.EOF Then Exit Function
        lng����ID = Val(nvl(rsTemp!����ID))
        strCardPassWord = nvl(rsTemp!����)
        If Val(nvl(rsTemp!״̬)) = 1 Then
            strErrMsg = "����Ϊ" & strCardNo & "�Ѿ�����ʧ!"
            If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        If Val(nvl(rsTemp!״̬)) = 2 Then
            strErrMsg = "����Ϊ" & strCardNo & "�Ѿ���ͣ��!"
             If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
             rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        GetPatiID = True
        Exit Function
    End If
    
     blnCard = True
    '����:47939
    Select Case UCase(strCardType)
    Case "IC��", "IC����"
        strWhere = "IC����=[2] "
    Case "���֤", "���֤��"
        strWhere = "���֤��=[2] "
    Case "��ϵ�����֤" '�����:51071
        strWhere = "��ϵ�����֤��=[2]"
    Case "ҽ����", "ҽ��֤��"
        strWhere = "ҽ����=[2] "
    Case "�ֻ���"
        strWhere = "�ֻ���=[2] "
    Case "�����"
        strWhere = "�����=[3] "
        str��ʶ�� = strCardNo
    '84247:���ϴ�,2015/4/24,סԺ�Ų��Ҳ���
    Case "סԺ��"
        strWhere = "a.����ID = (Select Nvl(Max(����ID),0) As ����ID From ������ҳ Where סԺ�� = [3]) "
        str��ʶ�� = strCardNo
    Case Else
        strWhere = "" & strCardType & "=[2] "
        blnCard = False
    End Select
    strSQL = "" & _
    "Select Rownum As ID, a.*" & vbNewLine & _
    "From (Select Decode(Nvl(max(a.��Ժ), 0), 1, '��', '') As ��Ժ, a.����id, max(a.����) As ����, max(a.�Ա�)as �Ա�, max(a.����) as ����, max(a.���֤��) as ���֤��," & vbNewLine & _
    "             max(a.Ic����) as IC����,max( a.�����)as �����, max(a.סԺ��)as סԺ��,max(a.�ֻ���)as �ֻ���,max( a.��������) as ��������, max(a.�����ص�) as �����ص�," & vbNewLine & _
    "             max(a.�ѱ�) as �ѱ�, max(a.ҽ�Ƹ��ʽ)as ҽ�Ƹ��ʽ, max(a.����) as ����,max(a.��ͥ��ַ) as ��ͥ��ַ, max(a.��ͥ�绰) as ��ͥ�绰," & vbNewLine & _
    "             max(a.��ϵ������) as ��ϵ������, max(a.��ϵ�˹�ϵ)as ��ϵ�˹�ϵ, max(a.��ϵ�˵绰) as ��ϵ�˵绰,max(a.��ϵ�����֤��) as ��ϵ�����֤��," & vbNewLine & _
    "             max(a.סԺ����)as סԺ����,max(a.����֤��) As ����id," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(����, 1, Nvl(b.Ԥ�����, 0), 0)), '99999999990.00')) As ����Ԥ�����," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(����, 1, 0, Nvl(b.Ԥ�����, 0))), '99999999990.00')) As סԺԤ�����" & IIf(lngShowCardNoTypeID <> 0, ",max(c.����) as ����", "") & vbNewLine & _
    "       From ������Ϣ A, ������� B" & IIf(lngShowCardNoTypeID <> 0, ",����ҽ�ƿ���Ϣ C", "") & vbNewLine & _
    "       Where a.ͣ��ʱ�� Is Null And a.����id = b.����id(+) And b.����(+) = 1 And " & IIf(lngShowCardNoTypeID <> 0, "a.����ID=c.����ID(+) And c.�����ID(+)=[5] And ", "") & strWhere & vbNewLine & _
    "       group by a.����ID" & vbNewLine & _
    "       Order By ����id" & IIf(lngShowCardNoTypeID <> 0, ", ����", "") & " ) A"
    Dim frmSel As New frmPatiSelect
    '52913
    '80886,Ƚ����,2014-12-18,������Ϊ"310D664700068D9E",ʹ��Val(strCardNo)�ᱨ"���"����
    If Not frmSel.ShowSelect(frmMain, cnOracle, glngSys, glngModul, objCtl, strSQL, "����ѡ��", "��ǰ���������������Ϣ,��ѡ��ָ���Ĳ���", True, blnShowMergePati, _
                             IIf(objCtl Is Nothing, False, True), "", "����ID,ID", rsTemp, blnUserCancel, Val(strCardType), strCardNo, Val(str��ʶ��), lng����ID, lngShowCardNoTypeID) Then
        If Not frmSel Is Nothing Then Unload frmSel
        Set frmSel = Nothing: Exit Function
    End If
    If Not frmSel Is Nothing Then Unload frmSel
    Set frmSel = Nothing
    
    If rsTemp Is Nothing Then GoTo GoClsObject:
    If rsTemp.State <> 1 Then GoTo GoClsObject:
    If rsTemp.EOF Then
        If blnCard Then
            'IC��,������ģ������
            GetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID)
        End If
        rsTemp.Close
        GoTo GoClsObject:
        Exit Function
    End If
    
    lng����ID = Val(nvl(rsTemp!����ID))
    strCardPassWord = nvl(rsTemp!����ID)
    GetPatiID = True
    Set objTemp = Nothing: Set objDatabase = Nothing
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        GoTo GoClsObject: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
GoClsObject:
    Set objTemp = Nothing: Set objDatabase = Nothing
    Set rsTemp = Nothing
End Function
 
Private Function GetCardNODencodeRule(ByVal lng�����ID As Long, _
    Optional bln���ѿ� As Boolean = False, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ID�Ĺ���
    '���:lng�����ID-�����ID
    '        bln���ѿ�-�Ƿ����ѿ�
    '����:�����Ŀ��ű������
    '����:���˺�
    '����:2011-06-22 11:01:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    If bln���ѿ� Then
        Set rsTemp = zlGet���ѿ��ӿ�(cnOracle)
        rsTemp.Filter = "���=" & lng�����ID
        If rsTemp.EOF Then GoTo GoEnd:
        GetCardNODencodeRule = nvl(rsTemp!�Ƿ�����)
        GoTo GoEnd:
    End If
    Set rsTemp = zlGetҽ�ƿ����(cnOracle)
    rsTemp.Filter = "ID=" & lng�����ID
    If rsTemp.EOF Then GoTo GoEnd:
    GetCardNODencodeRule = nvl(rsTemp!��������)
GoEnd:
    rsTemp.Filter = 0
End Function
Public Function GetCardNODencode(ByVal strCardNo As String, _
    Optional lng�����ID As Long = 0, _
    Optional strRule As String = "", Optional bln���ѿ� As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '���:lng�����ID-�����ID�����ѿ����,�������,����ҽ�ƿ��������ѿ����е�"��������"���Ƿ����Ľ��м���
    '       strRule-����:2-4��ʾ��2λ��4λ��*����,����-��,���ʾ�����λ��ʾΪ*
    '       strCardNo-����
    '����:
    '����:��**�Ŀ���,�������,���ؿ�
    '����:���˺�
    '����:2011-06-21 14:21:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant
    Dim strCardPassText As String, i As Long, j As Long
    If bln���ѿ� Then
        If Val(strRule) = 1 Then GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        If lng�����ID = 0 Then GetCardNODencode = strCardNo: Exit Function
        If Val(GetCardNODencodeRule(lng�����ID, True)) = 1 Then
            GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        Else
            GetCardNODencode = strCardNo: Exit Function
        End If
    End If
    If lng�����ID <> 0 And strRule = "" Then
        strCardPassText = GetCardNODencodeRule(lng�����ID)
    Else
        'ȡ�Ź���
        strCardPassText = strRule
    End If
    If strCardPassText = "" Then
       GetCardNODencode = strCardNo
    End If
    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '���λ��ʾ*
        i = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If i = 0 Then GetCardNODencode = strCardNo: Exit Function
        j = Len(strCardNo) - i: j = IIf(j < 0, 0, j)
        GetCardNODencode = Mid(strCardNo, 1, j) & String(i, "*")
        Exit Function
    End If
    i = Val(varPass(0)): j = Val(varPass(1))
    If i > Len(strCardNo) Then GetCardNODencode = strCardNo: Exit Function
    If j > Len(strCardNo) Then j = Len(strCardNo)
    If j < i Then j = i
   GetCardNODencode = Mid(strCardNo, 1, i - 1) & String(j - i + 1, "*") & Mid(strCardNo, j + 1)
End Function
Public Function InitInterFacel(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional bln���ѿ� As Boolean = False, Optional ByRef objPatiCard As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ָ�����ӿ�
    '���:lngCardTypeID-ָ�������
    '       bln���ѿ�-�Ƿ����ѿ�
    '����:��������True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    If Not objPatiCard Is Nothing Then
        If objPatiCard.�ӿ���� = lngCardTypeID And objPatiCard.���ѿ� = bln���ѿ� Then
            If Not objPatiCard.InitCompents Then
                If objPatiCard.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not objPatiCard.CardObject.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then                     '��ʼ������
                    Exit Function
                End If
                objPatiCard.InitCompents = True
            End If
            InitInterFacel = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln���ѿ�, "X", "K") & lngCardTypeID
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Set objPatiCard = objYlCardObjs(strKey)
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "����:" & lngCardTypeID & "δ�ҵ����" & IIf(bln���ѿ�, "���㿨", "ҽ�ƿ����") & "������,����!"
                Call WritLog("zlInitInterFacel", "", "����:" & lngCardTypeID & "δ�ҵ����" & IIf(bln���ѿ�, "���㿨", "ҽ�ƿ����") & "������,����!")
                Exit Function
            End If
            '���´���
            If zlCreatePatiCardObject(objCard, objCardObject) = False Then Exit Function
            '���Ӷ�Ӧ
           Set objPatiCard = objYlCardObjs.Add(objCardObject, objCard.���ƿ�, objCard.�ӿ����, objCard, bln���ѿ�, strKey)
    End If
    
    If objPatiCard Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "ע��:" & vbCrLf & "���ýӿ�(" & objCard.�ӿڱ��� & "-" & objCard.���� & ")����ʧ��,����!", vbInformation, gstrSysName
        Else
                MsgBox "ע��:" & vbCrLf & "���ýӿ�(" & lngCardTypeID & ")����ʧ��,����!", vbInformation, gstrSysName
        End If
        Exit Function
    End If

    Err = 0: On Error Resume Next
    Set objCard = objPatiCard.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "����:" & lngCardTypeID & "δ�ҵ�,����!" & vbCrLf & " ��ϸ�Ĵ�����Ϣ:" & Err.Description
        Call WritLog("clsPatiCard.zlInitInterFacel", "", "����:" & lngCardTypeID & "δ�ҵ�,����!" & vbCrLf & " ��ϸ�Ĵ�����Ϣ:" & Err.Description)
        Exit Function
    End If
    If Not objPatiCard.InitCompents Then
        If Not objPatiCard.CardObject.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         objPatiCard.InitCompents = True
    End If
    InitInterFacel = True
End Function

Public Function zlOnlyBrushCard(ByVal objEdit As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������(Ŀǰֻ֧���п�����ˢ��)
    '����:�Ƿ�ˢ��������,����true
    '����:���˺�
    '����:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean
    Dim strText As String
    'ˢ��ʱ����������ŵ��ɵ��÷�ȡ������
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    strText = objEdit.Text
    If objEdit.SelLength = Len(objEdit.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    sngNow = timer
    If objEdit.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
    End If
    If Not blnCard Then
        blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1
    End If
    
    If Not blnCard Then
        If gblnTestCardNo Then  '������Ϊˢ��
            If KeyAscii = 13 And Trim(objEdit.Text) <> "" Then
                 zlOnlyBrushCard = True
            End If
            Exit Function
        End If
        If KeyAscii <> 8 And KeyAscii <> 13 Then
            objEdit.Text = Chr(KeyAscii): objEdit.SelStart = Len(objEdit)
        Else
            objEdit.Text = ""
        End If
        If KeyAscii <> 13 Then
            KeyAscii = 0:
        End If
        Exit Function
    End If
    If KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 Then
        zlOnlyBrushCard = True
    End If
End Function
Public Function zlGetCardObj(ByVal frmMain As Object, ByVal lngCardTypeID As Long, _
    Optional bln���ѿ� As Boolean = False, _
    Optional ByRef objPatiCardObj As clsCardObject, _
    Optional ByRef blnNotParaCreateObject As Boolean = False, _
    Optional ByVal blnNotStartCreateObject As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������
    '���:lngCardTypeID-ָ�������
    '       bln���ѿ�-�Ƿ����ѿ�
    '       blnNotParaCreateObject-�����ݲ�����������
    '       blnNotStartCreateObject-Ϊtrueʱ��δ�������õ�ҲҪ�����ӿڶ���, _
    '                                                  ΪFalseʱ, ֻ�����������òŴ�������
    '����:��������True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    
    If Not objPatiCardObj Is Nothing Then
        If objPatiCardObj.�ӿ���� = lngCardTypeID And objPatiCardObj.���ѿ� = bln���ѿ� Then
            If Not objPatiCardObj.InitCompents Then
                If objPatiCardObj.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then                    '��ʼ������
                    Exit Function
                End If
                objPatiCardObj.InitCompents = True
            End If
            zlGetCardObj = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln���ѿ�, "X", "K") & lngCardTypeID
    '59760
    '����豸�Ƿ�����
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Set objPatiCardObj = objYlCardObjs(strKey)
    If Err <> 0 Or _
        blnNotParaCreateObject And objPatiCardObj.CardObject Is Nothing _
        Or blnNotStartCreateObject Then
          '����򲻸��ݲ�����������ʱ,����Ҫ���´�������
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "����:" & lngCardTypeID & "δ�ҵ����" & IIf(bln���ѿ�, "���㿨", "ҽ�ƿ����") & "������,����!"
                Call WritLog("zlInitInterFacel", "", "����:" & lngCardTypeID & "δ�ҵ����" & IIf(bln���ѿ�, "���㿨", "ҽ�ƿ����") & "������,����!")
                Exit Function
            End If
            'δ����ҲҪ��������
            If blnNotStartCreateObject Then objCard.���� = True
            '���´���
            If zlCreatePatiCardObject(objCard, objCardObject) = False Then Exit Function
            '���Ӷ�Ӧ
           Set objPatiCardObj = objYlCardObjs.Add(objCardObject, objCard.���ƿ�, objCard.�ӿ����, objCard, bln���ѿ�, strKey)
    End If
    
    If objPatiCardObj Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "ע��:" & vbCrLf & "���ýӿ�(" & objCard.�ӿڱ��� & "-" & objCard.���� & ")����ʧ��,����!", vbInformation, gstrSysName
        Else
                MsgBox "ע��:" & vbCrLf & "���ýӿ�(" & lngCardTypeID & ")����ʧ��,����!", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    Err = 0: On Error Resume Next
    Set objCard = objPatiCardObj.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "����:" & lngCardTypeID & "δ�ҵ�,����!" & vbCrLf & " ��ϸ�Ĵ�����Ϣ:" & Err.Description
        Call WritLog("clsPatiCard.zlInitInterFacel", "", "����:" & lngCardTypeID & "δ�ҵ�,����!" & vbCrLf & " ��ϸ�Ĵ�����Ϣ:" & Err.Description)
        Exit Function
    End If
    If Not objPatiCardObj.InitCompents Then
        If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         objPatiCardObj.InitCompents = True
    End If
    zlGetCardObj = True
End Function
Public Function zlSelectPayType(ByVal frmMain As Object, ByRef lngCardTypeID As Long, Optional blnNotTheeInterface As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��֧������
    '����:lngCardTypeID-�����ID
    '       blnNotTheeInterface-�����������ӿ�
    '����:���˺�
    '����:2012-06-11 14:11:20
    '����:ѡ��ɹ�,����true,���򷵻�False
    ''����:50120
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTypes As String, varData As Variant, strCardTypeIDs As String
    Dim i As Long, varTemp As Variant
    
    strTypes = GetAvailabilityCardType
    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    On Error GoTo errHandle
    strCardTypeIDs = ""
    varData = Split(strTypes, ";")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & "|||||||", "|")
        '�ݲ��������ѿ�
        If Val(varTemp(3)) <> 0 And Val(varTemp(5)) = 0 Then strCardTypeIDs = strCardTypeIDs & "," & Val(varTemp(3))
    Next
    If strCardTypeIDs = "" Then blnNotTheeInterface = True: Exit Function
    strCardTypeIDs = Mid(strCardTypeIDs, 2)
    If InStr(strCardTypeIDs, ",") = 0 Then
        'ֻ��һ�����ʱ
        lngCardTypeID = Val(strCardTypeIDs): zlSelectPayType = True: Exit Function
    End If
    '�������ѡ��һ��
    If Not frmSelectType.zlSelect(frmMain, strCardTypeIDs, lngCardTypeID, "֧����ʽѡ��") Then
      lngCardTypeID = 0: Exit Function
    End If
    If lngCardTypeID = 0 Then Exit Function
    zlSelectPayType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSelectWriteCardType(ByVal frmMain As Object, ByRef lngCardTypeID As Long, _
    Optional cnOracle As ADODB.Connection, Optional ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��(����/סԺ)д�����
    '����:lngCardTypeID-�����ID
    '����:���˺�
    '����:2012-2-12 15:21:22
    '����:ѡ��ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardTypeIDs As String
    Dim i As Long, objCard As clsCard
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    For i = 1 To objYlCardObjs.count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.���� And objCard.�Ƿ�д�� And objCard.���ѿ� = False Then
            strCardTypeIDs = strCardTypeIDs & "," & objCard.�ӿ����
        End If
    Next
    If strCardTypeIDs <> "" Then strCardTypeIDs = Mid(strCardTypeIDs, 2)
    strCardTypeIDs = ZLGetPatiCardFromCards(strCardTypeIDs, lng����ID)
    If strCardTypeIDs = "" Then Exit Function
    If InStr(strCardTypeIDs, ",") = 0 Then
        'ֻ��һ�����ʱ
        lngCardTypeID = Val(strCardTypeIDs): zlSelectWriteCardType = True: Exit Function
    End If
    '�������ѡ��һ��
    If Not frmSelectType.zlSelect(frmMain, strCardTypeIDs, lngCardTypeID, "ѡ��д�����", cnOracle) Then
      lngCardTypeID = 0: Exit Function
    End If
    If lngCardTypeID = 0 Then Exit Function
    zlSelectWriteCardType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZLGetPatiCardFromCards(ByVal strCardTypeIDs As String, ByVal lng����ID As Long) As String
    '�Ӹ���������м���ָ�����˳�����Ч���Ŀ����
    '���:
    '   strCardTypeIDs ��������𣬶���ö��ŷָ�
    '   lng����ID
    '���أ����ز��˳�����Ч���Ŀ���𣬶���ö��ŷָ�
    '����ţ�113121
    Dim strCards As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If strCardTypeIDs = "" Or lng����ID = 0 Then Exit Function
    strSQL = _
        "Select Distinct a.�����id" & vbNewLine & _
        "From ����ҽ�ƿ���Ϣ A" & vbNewLine & _
        "Where a.����id = [1] And Nvl(a.״̬, 0) = 0" & vbNewLine & _
        "      And �����id In (Select /*+cardinality(j,10)*/  j.Column_Value From Table(f_Num2list([2])) J)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatiCard", lng����ID, strCardTypeIDs)
    Do While Not rsTemp.EOF
        strCards = strCards & "," & Val(nvl(rsTemp!�����id))
        rsTemp.MoveNext
    Loop
    If strCards <> "" Then strCards = Mid(strCards, 2)
    ZLGetPatiCardFromCards = strCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardProperty(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As clsCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ƭ����
    '���:lngCardTypeID-�����ID
    '       bln���ѿ�-�Ƿ����ѿ�
    '����:objCard-������
    '����:����ָ�������ID�ģ�����true,���򷵻�False
    '����:���˺�
    '����:2014-01-17 16:07:24
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As clsCards, objTemp As clsCard
    On Error GoTo errHandle
    If zlGetCards_YL(objCards) = False Then Exit Function
    If objCards Is Nothing Then Exit Function
    For Each objTemp In objCards
        If objTemp.�ӿ���� = lngCardTypeID And objTemp.���ѿ� = bln���ѿ� Then
            Set objCard = objTemp
            zlGetCardProperty = True: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

