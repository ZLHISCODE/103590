Attribute VB_Name = "mdlDefine"
Option Explicit

Public gobjComLib As Object
Public gstrMessage As String                    '��Ϣ
Public gobjConn As ADODB.Connection             'HIS��DB���Ӷ���
Public gfrmOwner As Form                        '���������
Public glngSys As Long                          '��������ϵͳ��
Public glngModule As Long                       '��������ģ���
Public gstrDBUser As String                     'HIS��DB�û���
Public gstrRegHospital As String                'ע��ҽԺ����
Public gcolConn As Collection                   'clsConnect���󼯺�
Public gcolDevice As Collection                 'clsDevice���󼯺�
Public gobjSOAP As Object                       'MSSOAP����
Public gstrPrivs As String
'Public grsUpload As ADODB.Recordset
Public grsParam As ADODB.Recordset               '�������ݼ�

Public glngUserId As Long
Public glngDeptId As Long
Public gstrUserCode As String
Public gstrUserName As String
Public gstrUserAbbr As String
Public gstrDeptCode As String
Public gstrDeptName As String


Public gstrSQL As String

Public Const GINT_INTERFACE_MODULENO = 1348
Public Const GSTR_INTERFACE_NAME = "ҩ���Զ����ӿ�"
Public Const GSTR_SEPARAT = "|"
Public Const GSTR_SEPARAT_CHILD = ";"
Public Const GSTR_DEVICE_KEY = "D_"

'�Զ���ϵͳ��������
Public Enum enuLinkType
    DB
    WEBServices
    Directory
End Enum

'Ƕ��˵���
Public Enum enuMenuNo
    ҩƷ��Ϣ = 1
    ҩƷ���
    �豸����
    �ϴ�����
End Enum

Private Type Type_Params
    '�豸��Ӧ�Ĳ���
    int�������() As Integer                  '1-���2-סԺ
    int��ҩ��Ӧҵ��() As Integer              '1-�����շѣ�2-������ҩ��ҩ���ܣ�3-������ҩ��ҩ����
    bln���÷���֪ͨ() As Boolean              '1-����
    str��������() As String                   '��λ�ֱ��ʾ���������������˵���1��ʾѡ��0��ʾδѡ��
    strҩƷ����() As String                   'Null��ʾ����ҩƷ���ͣ������Ҫָ��ĳЩ���ͣ���ʽ��������,Ƭ��,��
    
    lngDeviceID() As Long                     '�豸ID
    lngStockID() As Long                      '�豸��Ӧ��ҩ��ID
    blnStart() As Boolean                     '�豸�Ƿ�����
End Type
Public gDeviceParams As Type_Params

Public Function GetDeviceID()

End Function

Public Function GetJudge_IsNeedUpload(ByVal lngModule As Boolean, ByVal bytType As Byte) As Boolean
'���ܣ��жϵ�ǰҵ�񻷽��Ƿ���Ҫ�ϴ�����
'������
'   lngModule��ģ���
'   bytType��
'       0: ҩƷ������Ϣ�ϴ�
'       1: ���ﴦ���ϴ� (��ҩ)
'       2: ���﷢ҩ֪ͨ (��ҩ)
'       3: סԺҩƷҽ���ϴ� (�䡢��ҩ)
'       4: ҩƷ����ϴ�

'    If grsParam Is Nothing Then
'        Call GetParam(0)
'    End If
    
    Select Case lngModule
        Case 1121   '�����շ�
            If bytType = 1 Then
                grsParam.Filter = "������='��ҩ��Ӧҵ��' And ����ֵ=1 "
                GetJudge_IsNeedUpload = Not grsParam.EOF
            End If
            Exit Function
        Case 1341   '������ҩ
            grsParam.Filter = "������='�������' And ����ֵ=1 "
            If grsParam.EOF Then
                Exit Function
            End If
            
            Select Case bytType
                Case 1
                    grsParam.Filter = "������='��ҩ��Ӧҵ��' And ����ֵ=1 "
                Case 2
                    grsParam.Filter = "������='���Ͷ�Ӧҵ��' And ����ֵ=1 "
                Case 0, 4
                    GetJudge_IsNeedUpload = True
                    Exit Function
                Case Else
                    Exit Function
            End Select
            
            GetJudge_IsNeedUpload = Not grsParam.EOF
            Exit Function
        Case 1342   '���ŷ�ҩ
            Select Case bytType
                Case 3
                    grsParam.Filter = "������='�������' And ����ֵ=2 "
                Case 0, 4
                    GetJudge_IsNeedUpload = True
                    Exit Function
                Case Else
                    Exit Function
            End Select
            GetJudge_IsNeedUpload = Not grsParam.EOF
            Exit Function
    End Select

End Function

Public Sub GetDeviceParam()
'���ܣ���ȡ�豸��Ӧ�Ĳ���ֵ������ŵ�����������
'������
'   lngDevicdID���豸ID
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    gstrSQL = "Select * From ҩ��ע���豸 Order by �豸ID "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam")
    
    Do While Not rsData.EOF
        ReDim Preserve gDeviceParams.lngDeviceID(UBound(gDeviceParams.lngDeviceID) + 1)
        ReDim Preserve gDeviceParams.lngStockID(UBound(gDeviceParams.lngStockID) + 1)
        ReDim Preserve gDeviceParams.blnStart(UBound(gDeviceParams.blnStart) + 1)
        
        ReDim Preserve gDeviceParams.int�������(UBound(gDeviceParams.int�������) + 1)
        ReDim Preserve gDeviceParams.int��ҩ��Ӧҵ��(UBound(gDeviceParams.int��ҩ��Ӧҵ��) + 1)
        ReDim Preserve gDeviceParams.bln���÷���֪ͨ(UBound(gDeviceParams.bln���÷���֪ͨ) + 1)
        ReDim Preserve gDeviceParams.str��������(UBound(gDeviceParams.str��������) + 1)
        ReDim Preserve gDeviceParams.strҩƷ����(UBound(gDeviceParams.strҩƷ����) + 1)
        
        gDeviceParams.lngDeviceID(UBound(gDeviceParams.lngDeviceID)) = Val(rsData!�豸ID)
        gDeviceParams.lngStockID(UBound(gDeviceParams.lngStockID)) = Val(rsData!����ID)
        gDeviceParams.blnStart(UBound(gDeviceParams.blnStart)) = (Val(NVL(rsData!����, 0)) = 1)
        
        rsData.MoveNext
    Loop
    
    gstrSQL = "Select a.�豸id, b.������, b.������, a.����ֵ, b.ȱʡֵ From ҩ���豸���� A, �Զ���ҩ���� B Where a.����id = b.Id Order by a.�豸id "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam")
    Do While Not rsData.EOF
        rsData.Filter = "������='�������'"
        If Not rsData.EOF Then gDeviceParams.int�������(rsData.AbsolutePosition - 1) = Val(NVL(rsData!����ֵ, rsData!ȱʡֵ))
        
        rsData.Filter = "������='��ҩ��Ӧҵ��'"
        If Not rsData.EOF Then gDeviceParams.int��ҩ��Ӧҵ��(rsData.AbsolutePosition - 1) = Val(NVL(rsData!����ֵ, rsData!ȱʡֵ))
        
        rsData.Filter = "������='���Ͷ�Ӧҵ��'"
        If Not rsData.EOF Then gDeviceParams.bln���÷���֪ͨ(rsData.AbsolutePosition - 1) = (Val(NVL(rsData!����ֵ, rsData!ȱʡֵ)) = 1)
        
        rsData.Filter = "������='��������'"
        If Not rsData.EOF Then gDeviceParams.str��������(rsData.AbsolutePosition - 1) = Val(NVL(rsData!����ֵ, rsData!ȱʡֵ))
        
        rsData.Filter = "������='ҩƷ����'"
        If Not rsData.EOF Then gDeviceParams.strҩƷ����(rsData.AbsolutePosition - 1) = Val(NVL(rsData!����ֵ, rsData!ȱʡֵ))
        
        rsData.MoveNext
    Loop
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetUserInfo()
    Dim strSQL As String
    Dim rsUser As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select R.*,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����, USER �û��� " & _
            " From �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա R" & _
            " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER and R.ȱʡ=1 " & _
            "       and (p.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null)"
    Set rsUser = gobjComLib.zldatabase.OpenSQLRecord(strSQL, "��ȡ�û���Ϣ")
    With rsUser
        If Not .EOF Then
            gstrDBUser = !�û���
            glngUserId = !��ԱID '��ǰ�û�id
            gstrUserCode = !��� '��ǰ�û�����
            gstrUserName = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            gstrUserAbbr = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            glngDeptId = !����ID '��ǰ�û�����id
            gstrDeptCode = !���ű��� '��ǰ�û�
            gstrDeptName = !�������� '��ǰ�û�
        Else
            gstrDBUser = ""
            glngUserId = 0 '��ǰ�û�id
            gstrUserCode = "" '��ǰ�û�����
            gstrUserName = "" '��ǰ�û�����
            gstrUserAbbr = "" '��ǰ�û�����
            glngDeptId = 0 '��ǰ�û�����id
            gstrDeptCode = "" '��ǰ�û�
            gstrDeptName = "" '��ǰ�û�
        End If
    End With
    Exit Function

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Public Function FindDeviceID(ByVal fldDeptID As Field, ByVal fldDrugType As Field, ByVal fldBill As Field, ByVal fldServiceObject As Field) As Long
'���ܣ���ȡע���豸ID
'������
'  fldDeptID��ҩ��ID
'  fldDrugType��ҩƷ����
'  fldBill����������
'  fldServiceObject���������
'���أ��豸ID

    Dim rsDevice As ADODB.Recordset
    Dim strTmp As String
    Dim strDrugType As String
    Dim strBill As String
    Dim strServiceObject As String
    Dim lngDeptID As Long, lngDeviceID As Long

    On Error GoTo errHandle

    'ҩ��ID
    lngDeptID = fldDeptID
    
    '�������
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldServiceObject))
    strServiceObject = IIf(strTmp = "", "0", IIf(strTmp = "����", "1", "2"))
    
    'ҩƷ����
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldDrugType))
    strDrugType = "%|" & IIf(strTmp = "", "????", strTmp) & "|%"
    
    '��������
    strTmp = Trim(gobjComLib.zlCommFun.NVL(fldBill))
    strBill = IIf(strTmp = "", "0", IIf(strTmp = "����", "1", IIf(strTmp = "����", "2", "3")))
    strBill = "%;" & strBill & ";%"
    
    gstrSQL = "Select Id " & _
              "From (Select a.Id, a.����, a.����, a.�ͺ�, a.����, Max(b.����) ������, " & _
              "        Max(Decode(d.������, 1, d.����ֵ, Null)) �������, " & _
              "        Max(Decode(d.������, 4, d.����ֵ, Null)) ��������, " & _
              "        Max(Decode(d.������, 5, d.����ֵ, Null)) ҩƷ����, " & _
              "        Max(Decode(d.������, 2, d.����ֵ, Null)) ��ҩҵ��, " & _
              "        Max(Decode(d.������, 3, d.����ֵ, Null)) ��ҩҵ��  " & _
              "      From ҩ��ע���豸 A, ҩ���豸���� B," & _
              "        (Select b.�豸id, b.����ֵ, a.������ From Zlparameters A, ҩ���豸���� B Where a.Id = b.����id) D " & _
              "         Where a.����id = b.Id And a.Id = d.�豸id(+) And a.����id = [1] " & _
              "      Group By a.Id, a.����, a.����, a.�ͺ�, a.����) A " & _
              "Where '|' || ҩƷ���� || '|' Like [2] and ������� = [3] "
    If strServiceObject = "2" Then
        '������������Ե������ͣ�ֻ�з�����סԺ���жϵ�������
        gstrSQL = gstrSQL & " and �������� like [4] "
    End If
    On Error GoTo errSQL
    Set rsDevice = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��ע���豸ID", lngDeptID, strDrugType, strServiceObject, strBill)
    On Error GoTo errHandle
    
    If rsDevice.EOF = False Then
        FindDeviceID = rsDevice!ID
    End If
    rsDevice.Close
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
    Exit Function

errSQL:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Function FindDevice(ByVal lngID As Long) As clsDevice
'���ܣ��ҵ��豸�������û���ҵ�����ʵ��һ��
'������
'   lngID���豸ID
'���أ�clsDevice����

    Dim strKey As String
    Dim i As Integer
    
    If lngID = 0 Then Exit Function
    
    If gcolDevice Is Nothing Then
        strKey = CreateDevice(lngID)
        If strKey <> "" Then Set FindDevice = gcolDevice(strKey)
    Else
        '���豸����
        If gcolDevice(GSTR_DEVICE_KEY & lngID) Is Nothing Then
            strKey = CreateDevice(lngID)
            If strKey <> "" Then Set FindDevice = gcolDevice(strKey)
        Else
            FindDevice = gcolDevice(GSTR_DEVICE_KEY & lngID)
        End If
    End If
    
    Exit Function
    
errHandle:
    Set FindDevice = Nothing
    gstrMessage = "δ�ҵ�����������ע���豸��"
End Function

Public Function CreateDevice(ByVal lngID As Long) As String
'���ܣ�ʵ���豸����
'������
'   lngDeptID���豸ID
'���أ��豸����Key
    Dim rsTmp As ADODB.Recordset
    Dim strKey As String
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select a.Id, a.����, a.����, a.�ͺ�, a.����id, a.����, Max(d.������) ������, Max(b.����) ������, " & _
              "    Max(Decode(d.������, 1, d.����ֵ, Null)) �������," & _
              "    Max(Decode(d.������, 4, d.����ֵ, Null)) ��������," & _
              "    Max(Decode(d.������, 5, d.����ֵ, Null)) ҩƷ����," & _
              "    Max(Decode(d.������, 2, d.����ֵ, Null)) ��ҩҵ��," & _
              "    Max(Decode(d.������, 3, d.����ֵ, Null)) ��ҩҵ�� " & _
              "From ҩ��ע���豸 A, ҩ���豸���� B, " & _
              "    (Select b.�豸id, b.����ֵ, a.������ From Zlparameters A, ҩ���豸���� B Where a.Id = b.����id) D " & _
              "Where a.����id = b.Id And a.Id = d.�豸id(+) And a.Id = [1] " & _
              "Group By a.Id, a.����, a.����, a.�ͺ�, a.����id, a.���� "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��ע���豸", lngID)
    If Not rsTmp.EOF Then
        strKey = GSTR_DEVICE_KEY & rsTmp!ID
        gcolDevice.Add New clsDevice, strKey
        With gcolDevice(strKey)
            .ID = rsTmp!ID
            .DeptID = rsTmp!DeptID
            .Link = gcolConn(rsTmp!������)
            .ServiceObject = gobjComLib.zlCommFun.NVL(rsTmp!�������, 0)
            .Bill = gobjComLib.zlCommFun.NVL(rsTmp!��������)
            .Enabled = gobjComLib.zlCommFun.NVL(rsTmp!����, 0) = 1
            .DrugType = gobjComLib.zlCommFun.NVL(rsTmp!ҩƷ����)
            .DispenseFunc = Val(gobjComLib.zlCommFun.NVL(rsTmp!��ҩҵ��))
            .DispensingFunc = Val(gobjComLib.zlCommFun.NVL(rsTmp!��ҩҵ��))
        End With
        CreateDevice = strKey
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    gstrMessage = "��δע���豸��Ϣ��ʵ���豸����ʧ�ܡ�"
End Function

Public Function TestURL(ByVal strURL As String) As Boolean
'���ܣ�����URL�Ƿ�����
'������
'  strURL��URL��ַ
'���أ�True���ӣ�Falseδ����
    Dim objSOAP As Object

    On Error Resume Next
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        gstrMessage = Err.Description
        Err.Clear
        On Error GoTo errSOAP
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
    End If
    
    '����
    objSOAP.MSSoapInit strURL
    If objSOAP.FaultCode <> "" Then
        gstrMessage = objSOAP.FaultString
        Set objSOAP = Nothing
    Else
        TestURL = True
        Set objSOAP = Nothing
    End If
    Exit Function
    
errSOAP:
    gstrMessage = Err.Description
End Function

Public Sub CreateWebServices(ByVal strURL As String, ByRef objWS As Object)
'���ܣ�����WebServices����
'������
'  strURL��
'  objWS��ʵ�ζ���

    On Error Resume Next
    Set objWS = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        gstrMessage = Err.Description
        Err.Clear
        On Error GoTo errSOAP
        Set objWS = CreateObject("MSSOAP.SoapClient")
    End If
    
    objWS.MSSoapInit strURL
    If objWS.FaultCode <> "" Then
        gstrMessage = objWS.FaultString
        Set objWS = Nothing
    End If
    Exit Sub
    
errSOAP:
    gstrMessage = Err.Description
    Set objWS = Nothing
End Sub

Public Function GetConnectStrEle(ByVal strConnect As String, ByVal bytType As Byte, ByVal strName As String) As String
'���ܣ���ȡ�������ݵ�Ԫ��ֵ
'������
'  strConnect����������
'  bytType����������
'  strName��Ҫ��ȡ��Ԫ����
'���أ�Ԫ��ֵ

    Dim arrEle As Variant
    Dim i As Integer

    Select Case bytType
        Case enuLinkType.WEBServices
            
            arrEle = Split(strConnect, GSTR_SEPARAT_CHILD)
            For i = LBound(arrEle) To UBound(arrEle)
                If UCase(strName) = Split(UCase(arrEle(i)), "=")(0) Then
                    GetConnectStrEle = Mid(arrEle(i), InStr(arrEle(i), "=") + 1)
                    Exit For
                End If
            Next
            Set arrEle = Nothing
    End Select
End Function

Public Function SetMenuItem(ByVal intFunc As Integer) As Boolean
'���ܣ����ù��ܲ˵���
'������
'  intFunc�����ܺ�
    
    Dim objMenuItem As Object
    
    On Error GoTo errHandle
    Load gfrmOwner.mnuDrugPackerItems(gfrmOwner.mnuDrugPackerItems.UBound + 1)
    Set objMenuItem = gfrmOwner.mnuDrugPackerItems(gfrmOwner.mnuDrugPackerItems.UBound)
    Select Case intFunc
        Case enuMenuNo.ҩƷ��Ϣ
            objMenuItem.Caption = "ҩƷ��Ϣ�ϴ�(&D)"
        Case enuMenuNo.ҩƷ���
            objMenuItem.Caption = "ҩƷ����ϴ�(&R)"
        Case enuMenuNo.�豸����
            objMenuItem.Caption = "�豸����(&S)"
        Case enuMenuNo.�ϴ�����
            objMenuItem.Caption = "�ϴ���������(&U)"
    End Select
    objMenuItem.Tag = intFunc
    SetMenuItem = True
    
errHandle:
    Set objMenuItem = Nothing
    If Err.Number <> 0 Then gstrMessage = "�Զ����ӿ�Ƕ��ʽ�˵�����ʧ�ܣ�"
End Function

Public Sub ShowMenuItem()
'���ܣ���ʾǶ��˵���
    With gfrmOwner
        .mnuDrugPacker.Visible = True
        .mnuDrugPackerItems(0).Visible = False
    End With
End Sub

'Public Function GetDeviceParam(ByVal lngDeviceID As Long, ByVal lngParamNO As Long) As String
''���ܣ���ȡָ���豸��ָ�������ŵĲ���ֵ
''������
''  lngDeviceID���豸ID
''  lngParamNO��������
''���أ��豸����ֵ
'
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    gstrSQL = "Select b.����ֵ From Zlparameters A, ҩ���豸���� B Where a.Id = b.����id And b.�豸id = [1] And a.������ = [2] "
'    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�豸�ķ������", lngDeviceID, lngParamNO)
'    If rsTmp.EOF = False Then
'        rstmp!����ֵ
'    End If
'    Exit Function
'
'errHandle:
'    If gobjComLib.ErrCenter = 1 Then Resume
'    gstrMessage = Err.Description
'End Function
 
Public Function GetHisRecord_DrugInf(ByVal lngDrugID As Long) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�ҩƷ������Ϣ
'������
'   lngDrugID��ҩƷID�����Ϊ0��ʾ����ҩƷ

    gstrSQL = "Select Decode(a.���, '5', '��ҩ', '6', '��ҩ', '��ҩ') As ����, e.����id, f.���� As ��������, g.ҩ��id As Ʒ��id, e.���� As Ʒ������," & vbNewLine & _
        " g.ҩƷid As ���id, h.ҩƷ���� As ����, e.����, a.���� As ͨ����, b.���� As ƴ������, c.���� As ��Ʒ��, d.���� As Ӣ����, a.���, e.���㵥λ As ������λ," & vbNewLine & _
        " g.����ϵ��, a.���㵥λ, g.���ﵥλ, g.�����װ, g.סԺ��λ, g.סԺ��װ, g.ҩ�ⵥλ, g.ҩ���װ, j.���� As �����̱��, a.���� As ����������, i.�ּ� As �ۼ�" & vbNewLine & _
        " From �շ���ĿĿ¼ A, �շ���Ŀ���� B, �շ���Ŀ���� C, �շ���Ŀ���� D, ������ĿĿ¼ E, ���Ʒ���Ŀ¼ F, ҩƷ��� G, ҩƷ���� H, �շѼ�Ŀ I, ҩƷ������ J" & vbNewLine & _
        " Where a.Id = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = 1 And a.Id = c.�շ�ϸĿid(+) And c.����(+) = 3 And c.����(+) = 1 And" & vbNewLine & _
        " a.Id = d.�շ�ϸĿid(+) And d.����(+) = 2 And a.Id = g.ҩƷid And g.ҩ��id = e.Id And e.����id = f.Id And g.ҩ��id = h.ҩ��id And" & vbNewLine & _
        " a.Id = i.�շ�ϸĿid And a.���� = j.����(+) And a.��� In ('5', '6', '7') And Sysdate Between i.ִ������ And" & vbNewLine & _
        " Nvl(i.��ֹ����, Sysdate) And a.����ʱ�� = Nvl(a.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))"
    If lngDrugID <> 0 Then
        gstrSQL = gstrSQL & " And A.id=[1] "
    End If
    gstrSQL = gstrSQL & " Order By Decode(a.���, '5', '��ҩ', '6', '��ҩ', '��ҩ'), a.Id"
    
    Set GetHisRecord_DrugInf = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_DrugInf", lngDrugID)

End Function

Public Function GetHisRecord_ReceipDetail(ByVal strKey As String) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�����ҩƷ��ϸ��Ϣ
'������
'   strKey������;�ⷿID;NO[|����;�ⷿID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int���� As Integer
    Dim lng�ⷿID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '�ֽ�Ϊ����
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '����ʽ�ַ����ֽⲢ�ֱ�ִ��SQL
        int���� = Split(arrKey(i), ";")(0)
        lng�ⷿID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select Distinct a.����, a.No, a.�������� As ����ʱ��, a.�ⷿid As ��ҩҩ��id, i.���� As ��ҩҩ��, a.���," & vbNewLine & _
            " Decode(b.���, '5', '��ҩ', '6', '��ҩ', '��ҩ') As ����, g.����id, k.���� As ��������, g.Id As Ʒ��id, g.���� As Ʒ������, j.ҩƷ����," & vbNewLine & _
            " a.ҩƷid, b.���� As ҩƷ����, b.���� As ҩƷ����, c.���� As ҩƷ��Ʒ��, h.���� As ҩƷӢ����, b.��� As ҩƷ���, g.���㵥λ As ������λ, d.����ϵ��," & vbNewLine & _
            " b.���㵥λ, d.���ﵥλ, d.�����װ, a.����, a.���� As ������, a.����, a.����, Nvl(a.����, 1) * a.ʵ������ / d.�����װ As ����," & vbNewLine & _
            " a.�ɱ��� * d.�����װ As �ɱ���, a.���ۼ� * d.�����װ As �ۼ�, e.Ӧ�ս��, e.ʵ�ս��, a.�÷� As ҩƷ�÷�, a.Ƶ��" & vbNewLine & _
            " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B, �շ���Ŀ���� C, ҩƷ��� D, ������ü�¼ E, ������ĿĿ¼ G, �շ���Ŀ���� H, ���ű� I, ҩƷ���� J, ���Ʒ���Ŀ¼ K" & vbNewLine & _
            " Where a.ҩƷid = b.Id And a.ҩƷid = c.�շ�ϸĿid(+) And c.����(+) = 3 And c.����(+) = 1 And a.ҩƷid = h.�շ�ϸĿid(+) And h.����(+) = 2 And" & vbNewLine & _
            " a.ҩƷid = d.ҩƷid And a.����id = e.Id And d.ҩ��id = g.Id And a.�ⷿid = i.Id And d.ҩ��id = j.ҩ��id And g.����id = k.Id And" & vbNewLine & _
            " a.���� = [1] And a.�ⷿid = [2] And a.No = [3] " & vbNewLine & _
            " Order By a.����, a.No, a.���"
        
        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipDetail", int����, lng�ⷿID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipDetail", int����, lng�ⷿID, strNO)
            
            '�����ݽ����ӵ���ʼ���ݼ���
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                'ע�⣺���SQL�������ӻ���٣���Ӧ����n�Ľ���ֵ��ĿǰSQLΪ34��
                For n = 0 To 33
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipDetail = rsData
End Function

Public Function GetHisRecord_ReceipList(ByVal strKey As String) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�������Ҫ��Ϣ
'������
'   strKey������;�ⷿID;NO[|����;�ⷿID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int���� As Integer
    Dim lng�ⷿID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '�ֽ�Ϊ����
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '����ʽ�ַ����ֽⲢ�ֱ�ִ��SQL
        int���� = Split(arrKey(i), ";")(0)
        lng�ⷿID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select a.����, a.No, Decode(a.��������, 1, '����', 2, '����', 3, '��һ', '4', '����', '5', '����', '��ͨ') As ��������, a.����id, a.��ҳid, a.����," & vbNewLine & _
            " c.�Ա�, c.����, c.��������, c.���, c.���￨��, c.�����, c.סԺ��, c.ҽ����, c.���֤��, c.Ic����, c.����, c.����, c.����, c.ҽ�Ƹ��ʽ As ҽ������," & vbNewLine & _
            " Sum(d.Ӧ�ս��) As �������, Sum(d.ʵ�ս��) As ʵ�ս��, a.�������� As ����ʱ��, d.��������id As ��������id, f.���� As ��������, d.������ As ����ҽ��," & vbNewLine & _
            " a.�ⷿid As ��ҩҩ��id, g.���� As ��ҩҩ��, Decode(a.���ȼ�, 1, '1', '2') As ���ȼ�, h.���� As ��ҩ���ڱ��, a.��ҩ����" & vbNewLine & _
            " From δ��ҩƷ��¼ A, ������Ϣ C, ������ü�¼ D, ҩƷ�շ���¼ E, ���ű� F, ���ű� G, ��ҩ���� H" & vbNewLine & _
            " Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id And e.����id = d.Id And d.��������id = f.Id And" & vbNewLine & _
            " a.�ⷿid = g.Id And a.��ҩ���� = h.����(+) And a.���� = [1] And a.�ⷿid = [2]  And a.No = [3] " & vbNewLine & _
            " Group By a.����, a.No, Decode(a.��������, 1, '����', 2, '����', 3, '��һ', '4', '����', '5', '����', '��ͨ'), a.����id, a.��ҳid, a.����, c.�Ա�," & vbNewLine & _
            " c.����, c.��������, c.���, c.���￨��, c.�����, c.סԺ��, c.ҽ����, c.���֤��, c.Ic����, c.����, c.����, c.����, c.ҽ�Ƹ��ʽ, a.��������, d.��������id," & vbNewLine & _
            " f.����, d.������, a.�ⷿid, g.����, Decode(a.���ȼ�, 1, '1', '2'), h.����, a.��ҩ����" & vbNewLine & _
            " Order By a.����, a.�ⷿid, Decode(a.���ȼ�, 1, '1', '2'), a.No, a.��������"
        
        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipList", int����, lng�ⷿID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipList", int����, lng�ⷿID, strNO)
            
            '�����ݽ����ӵ���ʼ���ݼ���
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                'ע�⣺���SQL�������ӻ���٣���Ӧ����n�Ľ���ֵ��ĿǰSQLΪ31��
                For n = 0 To 30
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipList = rsData
End Function

Public Function GetHisRecord_ReceipInf(ByVal strKey As String) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�������Ϣ��ҩƷ��ϸ
'������
'   strKey������;�ⷿID;NO[|����;�ⷿID;NO][|...]
    Dim rsData As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim int���� As Integer
    Dim lng�ⷿID As Long
    Dim strNO As String
    Dim i As Integer
    Dim n As Integer
    Dim arrKey As Variant
    
    '�ֽ�Ϊ����
    arrKey = Split(strKey, "|")
    For i = 0 To UBound(arrKey)
        If arrKey(i) = "" Or InStr(1, arrKey(i), ";") = 0 Then Exit For
        
        '����ʽ�ַ����ֽⲢ�ֱ�ִ��SQL
        int���� = Split(arrKey(i), ";")(0)
        lng�ⷿID = Split(arrKey(i), ";")(1)
        strNO = Split(arrKey(i), ";")(2)
        
        gstrSQL = "Select a.����, a.No, Decode(a.��������, 1, '����', 2, '����', 3, '��һ', '4', '����', '5', '����', '��ͨ') As ��������, a.����id, a.��ҳid, a.����," & vbNewLine & _
            " c.�Ա�, c.����, c.��������, c.���, c.���￨��, c.�����, c.סԺ��, c.ҽ����, c.���֤��, c.Ic����, c.����, c.����, c.����, c.ҽ�Ƹ��ʽ As ҽ������," & vbNewLine & _
            " a.�������� As ����ʱ��, d.��������id As ��������id, f.���� As ��������, d.������ As ����ҽ��, a.�ⷿid As ��ҩҩ��id, g.���� As ��ҩҩ��," & vbNewLine & _
            " Decode(a.���ȼ�, 1, '1', '2') As ���ȼ�, h.���� As ��ҩ���ڱ��, a.��ҩ����, e.���, Decode(i.���, '5', '��ҩ', '6', '��ҩ', '��ҩ') As ����," & vbNewLine & _
            " l.����id, o.���� As ��������, l.Id As Ʒ��id, l.���� As Ʒ������, n.ҩƷ����, e.ҩƷid, i.���� As ҩƷ����, i.���� As ҩƷ����, j.���� As ҩƷ��Ʒ��," & vbNewLine & _
            " m.���� As ҩƷӢ����, i.��� As ҩƷ���, l.���㵥λ As ������λ, k.����ϵ��, i.���㵥λ, k.���ﵥλ, k.�����װ, e.����, e.���� As ������, e.����, e.����," & vbNewLine & _
            " Nvl(e.����, 1) * e.ʵ������ / k.�����װ As ����, e.�ɱ��� * k.�����װ As �ɱ���, e.���ۼ� * k.�����װ As �ۼ�, d.Ӧ�ս��, d.ʵ�ս��, e.�÷� As ҩƷ�÷�," & vbNewLine & _
            " e.Ƶ��" & vbNewLine & _
            " From δ��ҩƷ��¼ A, ������Ϣ C, ������ü�¼ D, ҩƷ�շ���¼ E, ���ű� F, ���ű� G, ��ҩ���� H, �շ���ĿĿ¼ I, �շ���Ŀ���� J, ҩƷ��� K, ������ĿĿ¼ L, �շ���Ŀ���� M, ҩƷ���� N," & vbNewLine & _
            " ���Ʒ���Ŀ¼ O" & vbNewLine & _
            " Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id And e.����id = d.Id And d.��������id = f.Id And" & vbNewLine & _
            " a.�ⷿid = g.Id And a.��ҩ���� = h.����(+) And e.ҩƷid = i.Id And e.ҩƷid = j.�շ�ϸĿid(+) And j.����(+) = 3 And j.����(+) = 1 And" & vbNewLine & _
            " e.ҩƷid = m.�շ�ϸĿid(+) And m.����(+) = 2 And e.ҩƷid = k.ҩƷid And k.ҩ��id = l.Id And k.ҩ��id = n.ҩ��id And l.����id = o.Id And" & vbNewLine & _
            " a.���� = [1] And a.�ⷿid = [2] And a.No = [3] " & vbNewLine & _
            " Order By a.����, a.�ⷿid, a.No, e.���"

        If i = 0 Then
            Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipInf", int����, lng�ⷿID, strNO)
        Else
            Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_ReceipInf", int����, lng�ⷿID, strNO)
            
            '�����ݽ����ӵ���ʼ���ݼ���
            Do While Not rsTmp.EOF
                rsData.AddNew
                
                'ע�⣺���SQL�������ӻ���٣���Ӧ����n�Ľ���ֵ��ĿǰSQLΪ58��
                For n = 0 To 57
                    rsData.Fields(n).Value = rsTmp.Fields(n).Value
                Next
                
                rsData.Update
                
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    Set GetHisRecord_ReceipInf = rsData
End Function

Public Function GetHisRecord_AdviceInf(ByVal strKey As String) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�ҽ����Ϣ��ҩƷ��ϸ
'������
'   strKey��ҩƷID������ʽΪ"ҩƷID,ҩƷID..."

    gstrSQL = "Select /*+ rule*/ a.����id, a.��ʶ�� As סԺ��, a.����, a.����, a.�Ա�, a.����, q.��������, q.���, q.���￨��, q.ҽ����, q.���֤��, q.Ic����, q.����, q.����, q.����," & vbNewLine & _
        " a.��������id As ��������id, r.���� As �������ű���, r.���� As ������������, a.���˿���id, s.���� As ���˿��ұ���, s.���� As ���˿�������, a.���˲���id," & vbNewLine & _
        " f.���� As ���˲�������, f.���� As ���˲�������, b.�Է�����id As ��ҩ����id, t.���� As ��ҩ���ű���, t.���� As ��ҩ��������," & vbNewLine & _
        " Decode(d.ҽ����Ч, 1, '����', '��ʱ') As ҽ������, a.������ As ����ҽ��, c.����ʱ�� As ҽ������ʱ��, c.�״�ʱ��, c.ĩ��ʱ��, d.ִ��Ƶ��, d.Ƶ�ʴ���, d.Ƶ�ʼ��," & vbNewLine & _
        " d.�����λ, d.ִ��ʱ�䷽��, d.ҽ������, b.�÷� As ҩƷ�÷�, Decode(g.���, '5', '��ҩ', '6', '��ҩ', '��ҩ') As ����, h.����id, m.���� As ��������," & vbNewLine & _
        " i.ҩ��id As Ʒ��id, h.���� As Ʒ������, l.ҩƷ����, b.ҩƷid, g.���� As ҩƷ����, g.���� As ҩƷ����, n.���� As ҩƷ��Ʒ��, o.���� As ҩƷӢ����, g.���," & vbNewLine & _
        " b.���� As ������, b.����, b.����, i.����ϵ��, h.���㵥λ As ������λ, g.���㵥λ, i.סԺ��λ, i.סԺ��װ, b.����," & vbNewLine & _
        " Nvl(b.����, 1) * b.ʵ������ / i.סԺ��װ As ����, b.�ɱ��� * i.סԺ��װ As �ɱ���, b.���ۼ� * i.סԺ��װ As �ۼ�, b.���۽�� As ���, b.Id As �շ�id," & vbNewLine & _
        " b.�ⷿid As ��ҩҩ��id, u.���� As ��ҩҩ�����, u.���� As ��ҩҩ��, b.��������, b.�������" & vbNewLine & _
        " From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ������ C, ����ҽ����¼ D, ���ű� F, �շ���ĿĿ¼ G, ������ĿĿ¼ H, ҩƷ��� I, ҩƷ���� L, ���Ʒ���Ŀ¼ M, �շ���Ŀ���� N, �շ���Ŀ���� O," & vbNewLine & _
        " ������Ϣ Q, ���ű� R, ���ű� S, ���ű� T, ���ű� U , Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V " & vbNewLine & _
        " Where a.Id = b.����id And a.ҽ����� = c.ҽ��id And c.ҽ��id = d.Id And b.No = c.No And a.���˲���id = f.Id And b.ҩƷid = g.Id And" & vbNewLine & _
        " h.Id = i.ҩ��id And b.ҩƷid = i.ҩƷid And i.ҩ��id = l.ҩ��id And h.����id = m.Id And g.Id = n.�շ�ϸĿid(+) And n.����(+) = 3 And" & vbNewLine & _
        " n.����(+) = 1 And g.Id = o.�շ�ϸĿid(+) And o.����(+) = 2 And a.����id = q.����id And a.��������id = r.Id And a.���˿���id = s.Id And" & vbNewLine & _
        " b.�Է�����id = t.Id And b.�ⷿid = u.Id And b.Id = v.Column_Value "
    Set GetHisRecord_AdviceInf = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_AdviceInf", strKey)

End Function

Public Function GetHisRecord_DrugStock(ByVal lngStockID As Long) As ADODB.Recordset
'���ܣ���ȡHIS�˻������ݣ�ҩƷ�����Ϣ
'������
'   lngStockID���ⷿID

    gstrSQL = "Select Decode(a.���, '5', '��ҩ', '6', '��ҩ', '��ҩ') As ����, e.����id, f.���� As ��������, g.ҩ��id As Ʒ��id, e.���� As Ʒ������," & vbNewLine & _
        " g.ҩƷid As ���id, h.ҩƷ���� As ����, e.����, a.���� As ͨ����, b.���� As ƴ������, c.���� As ��Ʒ��, d.���� As Ӣ����, a.���," & vbNewLine & _
        " Decode(a.�Ƿ���, 1, 'ʱ��', '����') As �۸�����, e.���㵥λ As ������λ, g.����ϵ��, a.���㵥λ, g.���ﵥλ, g.�����װ, g.סԺ��λ, g.סԺ��װ, g.ҩ�ⵥλ," & vbNewLine & _
        " g.ҩ���װ, i.�ּ� As �ۼ�, k.����, k.Ч��, k.��������, k.ʵ������, k.ʵ�ʽ�� As ʵ�ʽ��, k.ʵ�ʲ�� As ʵ�ʲ��, l.���� As ��Ӧ��, k.�ϴβɹ��� As �ɹ���," & vbNewLine & _
        " k.�ϴ����� As ����, k.�ϴ��������� As ��������, k.�ϴβ��� As ����, k.��׼�ĺ�, k.ƽ���ɱ���" & vbNewLine & _
        " From �շ���ĿĿ¼ A, �շ���Ŀ���� B, �շ���Ŀ���� C, �շ���Ŀ���� D, ������ĿĿ¼ E, ���Ʒ���Ŀ¼ F, ҩƷ��� G, ҩƷ���� H, �շѼ�Ŀ I, ҩƷ��� K, ��Ӧ�� L" & vbNewLine & _
        " Where a.Id = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = 1 And a.Id = c.�շ�ϸĿid(+) And c.����(+) = 3 And c.����(+) = 1 And" & vbNewLine & _
        " a.Id = d.�շ�ϸĿid(+) And d.����(+) = 2 And a.Id = g.ҩƷid And g.ҩ��id = e.Id And e.����id = f.Id And g.ҩ��id = h.ҩ��id And" & vbNewLine & _
        " a.Id = i.�շ�ϸĿid And a.��� In ('5', '6', '7') And Sysdate Between i.ִ������ And Nvl(i.��ֹ����, Sysdate) And" & vbNewLine & _
        " a.����ʱ�� = Nvl(a.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And g.ҩƷid = k.ҩƷid And k.���� = 1 And" & vbNewLine & _
        " k.�ϴι�Ӧ��id = l.Id(+) And �ⷿid = [1] " & vbNewLine & _
        " Order By Decode(a.���, '5', '��ҩ', '6', '��ҩ', '��ҩ'), a.Id, k.����"
    Set GetHisRecord_DrugStock = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetHisRecord_DrugStock", lngStockID)

End Function
