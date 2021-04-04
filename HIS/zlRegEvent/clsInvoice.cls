VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'����:Ʊ�ݿ���

'����:LNC
'����:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mfrmMain As Object
Private mblnSharedInvoice As Boolean
Private mblnStartFactUseType As Boolean
Private mbln�ϸ���� As Boolean
Private mblnHisMgrInvoice As Boolean 'his����Ʊ�ݣ����õ���Ʊ��ʱ��Ч
Private mblnPrint As Boolean
Private mbyt���� As Byte
Private mbytInvoiceLen As Byte
Private mbytInvoiceType As Byte '��Ʊ��ӡ��ʽ
Private mbytInvoiceFormat As Byte '��Ʊ��ӡ��ʽ
Private mlngModule As Long
Private mlng����ID As Long
Private mlng��������ID As Long
Private mstrUseType As String
Private mstr���� As String
Private mstrInvoice As String

Public Property Let strInvoice(ByVal varValue As String)
    mstrInvoice = varValue
End Property

Public Property Get strInvoice() As String
    strInvoice = mstrInvoice
End Property

Public Property Get strUseType() As String
    strUseType = mstrUseType
End Property

Public Property Let blnPrint(ByVal varValue As Boolean)
    mblnPrint = varValue
End Property

Public Property Get blnPrint() As Boolean
    blnPrint = mblnPrint
End Property

Public Property Let blnHisMgrInvoice(ByVal varValue As Boolean)
    mblnHisMgrInvoice = varValue
End Property

Public Property Get blnHisMgrInvoice() As Boolean
    blnHisMgrInvoice = mblnHisMgrInvoice
End Property

Public Property Get bytPrintType() As Byte
    bytPrintType = mbytInvoiceType
End Property

Public Property Get bytPrintFormat() As Byte
    bytPrintFormat = mbytInvoiceFormat
End Property

Public Property Get blnSharedInvoice() As Boolean
    blnSharedInvoice = mblnSharedInvoice
End Property

Public Property Get blnStartFactUseType() As Boolean
    blnStartFactUseType = mblnStartFactUseType
End Property

Public Property Get bln�ϸ����() As Boolean
    bln�ϸ���� = mbln�ϸ����
End Property

Public Property Get lng��������ID() As Long
    lng��������ID = mlng��������ID
End Property

Public Property Get lng����ID() As Long
    lng����ID = mlng����ID
End Property


Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt���� As Byte = 4, Optional ByVal lngModule As Long = 1111) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ʊ�ݶ���
    '���:
    '    byt���ϣ�4-�Һ�;5-����
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    1.ʹ�ñ�����ǰ,�����ȵ��ñ��ӿڽ��г�ʼ��
    '    2.��ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ�������)
    '    3.ģ��ţ�1111=����Һ�,1121=�����շ�,1107=ҽ�ƿ����Ź���;
    '������:
    '   �Һ�?�Һ�Ʊ������?
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule
    mbyt���� = byt����
    Select Case byt����
        Case 4
            mstr���� = "�Һ�"
        Case 5
            mstr���� = "�շ�"
    End Select
    
    If zlInitPara(byt����) = False Then Exit Function
    zlInitialize = True
End Function

Public Function zlTerminate() As Boolean
    Set mfrmMain = Nothing
End Function

Private Function zlInitPara(ByVal byt���� As Byte) As Boolean
    Dim strValue As String
    
    '�Һŷ���Ҳʹ�ùҺŴ�ӡ��ʽ����
    mbytInvoiceType = Val(zlDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, mlngModule))
    If byt���� = 4 Then
        '�Һ�ʹ���շ�Ʊ��
        mblnSharedInvoice = zlDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121) = "1"
        
    ElseIf byt���� = 5 Then
        'ҽ�ƿ�ʹ���շ�Ʊ��
        mblnSharedInvoice = zlDatabase.GetPara("����ʹ�������շ�ҽ���վ�", glngSys, 1107) = "1"
        If Not mblnSharedInvoice Then zlInitPara = True: Exit Function
        
        mbytInvoiceFormat = Val(Split(zlDatabase.GetPara("ҽ�ƿ��վݸ�ʽ", glngSys, 1107) & "|", "|")(0))
    End If
    
    '���ع��ùҺ�����ID
    If mblnSharedInvoice Then
        mlng��������ID = Val(zlDatabase.GetPara("�����շ�Ʊ������", glngSys, 1121, ""))
    Else
        mlng��������ID = Val(zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, mlngModule, ""))
    End If
    If mlng��������ID > 0 Then
        If Not ExistShareBill(mlng��������ID, IIf(mblnSharedInvoice, 1, 4)) Then
            Call SaveCurInvoice("0")
            mlng��������ID = 0
        End If
    End If
    '�Ƿ������շ�Ʊ��ʹ�����
    If mblnSharedInvoice Then
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    
    'Ʊ���Ƿ��ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mbln�ϸ���� = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    'Ʊ�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    mbytInvoiceLen = Val(Split(strValue, "|")(IIf(gblnSharedInvoice, 0, 3)))
    
    zlInitPara = True
End Function

Public Function GetUseType(ByVal lng����id As Long, ByVal intInsure As Integer) As String
    '����:��ȡ�շ�Ʊ�ݵ�ʹ�����
    Dim strUseType As String
    Dim varData As Variant, varTemp As Variant
    mstrUseType = "": mlng��������ID = 0: mlng����ID = 0
    If mblnStartFactUseType Then
        mstrUseType = zl_GetInvoiceUserType(lng����id, 0, intInsure)
        mlng��������ID = zl_GetInvoiceShareID(1121, strUseType)
        If mbytInvoiceType <> 0 Then
            Call zlGetInvoiceGroupUseID(mlng����ID)
        End If
    End If
    GetUseType = mstrUseType
End Function

Public Function SaveCurInvoice(ByVal strInvoice As String) As Boolean
    '���ܣ����浱ǰʹ�õ�Ʊ�ݺ�
    If mbln�ϸ���� Then SaveCurInvoice = True: Exit Function
    If mbyt���� = 4 Then
        If mblnSharedInvoice Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strInvoice, glngSys, 1121
        Else
            zlDatabase.SetPara "��ǰ�Һ�Ʊ�ݺ�", strInvoice, glngSys, mlngModule
        End If
    Else
        zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strInvoice, glngSys, 1121
    End If
    SaveCurInvoice = True
End Function

Public Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbyt���� <> 4 Then Exit Function
    mlng����ID = GetInvoiceGroupID(IIf(mblnSharedInvoice, 1, 4), intNum, lng����ID, mlng��������ID, strInvoiceNO, IIf(mblnStartFactUseType, mstrUseType, ""))
    If mlng����ID <= 0 Then
        Select Case mlng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mstrUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mstrUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
                On Error Resume Next
                zlControl.ControlSetFocus mfrmMain.txtFact
                Exit Function
        End Select
    End If
    lng����ID = mlng����ID
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckUsedBill(ByVal lng����ID As Long, Optional ByVal strInvoice As String) As Long
    '��ȡ����ID�������ʹ���շ�Ʊ�ݣ��ڵ���ǰ����Ҫ�ȵ���GetUseType()
    If mbyt���� = 4 Then
        mlng����ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(lng����ID > 0, lng����ID, mlng��������ID), strInvoice, mstrUseType)
    ElseIf mbyt���� = 5 Then
        If Not mblnSharedInvoice Then Exit Function
        mlng����ID = CheckUsedBill(1, IIf(lng����ID > 0, lng����ID, mlng��������ID), strInvoice, mstrUseType)
    Else
        mlng����ID = 0
    End If
    zlCheckUsedBill = mlng����ID
End Function

Public Function zlCheckBillRepeat(ByVal lng����ID As Long, ByVal strInvoice As String) As Boolean
    '��鲢����
    If mbyt���� = 4 Then
        zlCheckBillRepeat = CheckBillRepeat(lng����ID, IIf(mblnSharedInvoice, 1, 4), strInvoice)
    ElseIf mbyt���� = 5 Then
        If Not mblnSharedInvoice Then Exit Function
        zlCheckBillRepeat = CheckBillRepeat(lng����ID, 1, strInvoice)
    End If
End Function

Public Function zlCheckInvoiceIsValid(ByRef lng����ID As Long, Optional ByRef strInvoice As String, Optional ByVal blnAutoInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���Ʊ����Ч��
    ' ��� : strInvoice��ʵ��Ʊ��
    '        blnAutoInput:Ʊ���Ƿ��Զ���ȡ��
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/8 19:25
    '---------------------------------------------------------------------------------------
    If Not mblnPrint Then zlCheckInvoiceIsValid = True: Exit Function
    
    If lng����ID = 0 Then lng����ID = mlng����ID
    If mbln�ϸ���� Then
        If strInvoice = "" Then
            If mbyt���� = 4 Then
                MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                Exit Function
            Else
                '�Զ����������շ�Ʊ��
                strInvoice = zlGetNextInvoice(lng����ID)
            End If
        End If
InvoiceHandle:
        lng����ID = zlCheckUsedBill(lng����ID, strInvoice)
        If lng����ID <= 0 Then
            Select Case lng����ID
            Case 0    '����ʧ��
            Case -1
                MsgBox "��û�����ú͹��õ�" & mstr���� & "Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -3
                MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        
        '�����������,Ʊ���Ƿ�����
        If zlCheckBillRepeat(lng����ID, strInvoice) Then
            If Not blnAutoInput Then
                MsgBox "Ʊ�ݺ�""" & strInvoice & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                Exit Function
            Else
                strInvoice = zlGetNextInvoice(lng����ID)
                If strInvoice = "" Then Exit Function
                
                MsgBox "��ǰƱ�ݺ��Ѿ���ʹ�ã������»�ȡƱ�ݺ�:" & strInvoice, vbInformation, gstrSysName
                GoTo InvoiceHandle
            End If
        End If
    Else
        If Len(strInvoice) <> mbytInvoiceLen And strInvoice <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & mbytInvoiceLen & " λ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mlng����ID = lng����ID
    zlCheckInvoiceIsValid = True
End Function

Public Function zlGetNextInvoice(ByRef lng����ID As Long) As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��һ����Ч��Ʊ��
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/4/8 19:37
    '---------------------------------------------------------------------------------------
    Dim strCurInvoice As String
    
    If mbln�ϸ���� Then
        lng����ID = zlCheckUsedBill(lng����ID)
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�" & mstr���� & "Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        '�ϸ�ȡ��һ������
        zlGetNextInvoice = GetNextBill(lng����ID)
    Else
        '��ɢ��ȡ��һ������
        If mblnSharedInvoice Then
            strCurInvoice = zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
        Else
            strCurInvoice = zlDatabase.GetPara("��ǰ�Һ�Ʊ�ݺ�", glngSys, mlngModule)
        End If
        zlGetNextInvoice = zlStr.Increase(strCurInvoice)
    End If
End Function