Attribute VB_Name = "mdl��Ϫũҽ"
Option Explicit
Public Declare Sub CXNY_SetRemoteServerAddr Lib "JFNetLib.dll" Alias "SetRemoteServerAddr" _
    (ByVal lngPort As Long, ByVal strIP As String)
Public Declare Function CXNY_SendRequestPack Lib "JFNetLib.dll" Alias "SendRequestPack" _
    (ByVal strSend As String, ByVal lngSend As Long, ByVal strReceive As String, lngReceive As Long, ByVal lngWaitSecs As Long) As Long

'ҵ�����幦�ܺ���
Public Const gstrFunc��Ϫũҽ_GetServerTime As String = "H888"     'ȡ������ʱ��
Public Const gstrFunc��Ϫũҽ_GetHospitalInfo As String = "H301"   'ȡҽԺ���Ƽ�����
Public Const gstrFunc��Ϫũҽ_GetPersonalInfo As String = "H302"   'ȡ������Ϣ
Public Const gstrFunc��Ϫũҽ_OutRegist As String = "H101"         '����Ǽ�
Public Const gstrFunc��Ϫũҽ_InRegist As String = "H201"          '��Ժ�Ǽ�
Public Const gstrFunc��Ϫũҽ_InRegistCancel As String = "H401"    '��Ժ�Ǽ�ȡ��
Public Const gstrFunc��Ϫũҽ_UploadDetail As String = "H303"      '�ϴ�������ϸ
Public Const gstrFunc��Ϫũҽ_InBalance As String = "H202"         'סԺ����
Public Const gstrFunc��Ϫũҽ_OutBalance As String = "H102"        '�������
Public Const gstrFunc��Ϫũҽ_BalanceCancel As String = "H103"     '��������
Public Const gstrFunc��Ϫũҽ_ModifyInfo As String = "H606"        '�޸�סԺ��Ϣ
Public Const gstrFunc��Ϫũҽ_DelAllDetail As String = "H607"      'ɾ��סԺ�����������ϴ�������ϸ
Public Const gstrFunc��Ϫũҽ_DelAllDetail1 As String = "H608"      'ɾ�������������ϴ�������ϸ
'���ڶ��ʵĹ��ܺ���
Public Const gstrFunc��Ϫũҽ_OutQuery As String = "H601"          '�˶�����
Public Const gstrFunc��Ϫũҽ_InQuery As String = "H602"           '�˶�סԺ
Public Const gstrFunc��Ϫũҽ_AllowQuery As String = "H603"        '������ѯ
Public Const gstrFunc��Ϫũҽ_Exception_Out As String = "H604"     '����ʱ������ҵ��ȡ��
Public Const gstrFunc��Ϫũҽ_Exception_In As String = "H605"      '����ʱ��סԺҵ��ȡ��

Private Const mstrAmountFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrPriceFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrMoneyFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrDateFormat As String = "yyyy-MM-dd HH:mm:ss"
Private Const mstrSplit As String = "&"
Private mblnInit As Boolean                                         '�Ƿ�������ʼ��

Private Type ComInfo_��Ϫũҽ
    ҽԺ���� As String
    ҽԺ���� As String
    ҵ������ As String
    ҽ��֤�� As String
    ���˱�� As String
    ������ˮ�� As String
    ������ˮ�� As String
    �ܷ��� As Currency                      'HIS
    �ܷ���_���� As Currency                 '���ĵķ����ܶ�
    ���㴮 As String
    ���㴮�� As String
    �������� As String
    �������� As String
    
End Type
Public gComInfo_��Ϫũҽ As ComInfo_��Ϫũҽ

'06-03-25Ӧ�����޸�
Public g����ũ������ As Currency    '��������ũ�����Ա����Ľ��
Public g������ˮ�� As String
Public g������ʱ��־ As String      '�Ƿ�����
'06-03-25



Private mstrFunc As String              '���ܺ�
Private mstrInput As String             '���봮
Private mlngInput As Long               '���봮����
Private mstrOutput As String            '�����
Private mlngOutput As Long              '���������
Private mlngReturn As Long              '�ȴ�����>=30
Public gstrOutput_��Ϫũҽ As String

'�����ô���----------------------------------------------
'Private gobjCXXNY As New clsT_CXXNY

Public Function ҽ������_��Ϫũҽ() As Boolean
    ҽ������_��Ϫũҽ = frmSet��Ϫũҽ.ShowME
End Function

Public Sub ���ýӿ�_׼��_��Ϫũҽ(ByVal strFunc As String, ByVal StrInput As String)
    mstrFunc = strFunc
    mstrOutput = String(2000, " ")
    mlngOutput = 2000
    mstrInput = "exchcode=" & mstrFunc & mstrSplit & StrInput
    mlngInput = LenB(StrConv(mstrInput, vbFromUnicode))
End Sub

Public Function ���ýӿ�_��Ϫũҽ() As Boolean
    Dim strMsg As String
    Dim arrReturn
    Dim blnSuccess As Boolean
    
'    mlngReturn = gobjCXXNY.CXNY_SendRequestPack(mstrInput, mlngInput, mstrOutput, mlngOutput, 30)
    mlngReturn = CXNY_SendRequestPack(mstrInput, mlngInput, mstrOutput, mlngOutput, 30)
    Select Case mlngReturn
    Case 0
        blnSuccess = True
    Case 1
        strMsg = "����ҽԺǰ�û�������ʧ��"
    Case 2
        strMsg = "��Զ�̷�������������ʧ��"
    Case 3
        strMsg = "���շ���ֵ���ʧ��"
    Case 4
        strMsg = "��̬���ӿⲻ����"
    Case Else
        strMsg = "����δ֪����"
    End Select
    
    arrReturn = Split(mstrOutput, "&")
    If blnSuccess Then
        blnSuccess = (Val(Split(arrReturn(0), "=")(1)) = 0)
        If blnSuccess = False Then strMsg = Split(arrReturn(1), "=")(1)
    End If
    
    If blnSuccess = False Then
        MsgBox strMsg & vbCrLf & "���ܣ�" & mstrFunc & "|����ţ�" & mlngReturn, vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrOutput_��Ϫũҽ = mstrOutput
    ���ýӿ�_��Ϫũҽ = True
End Function

Public Function ��ݱ�ʶ_��Ϫũҽ(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim arrReturn
    Dim StrInput As String
    Dim intVerify As Integer
    Dim strDiseaseCode As String            '��������
    Dim strIdentify As String
    Dim strRegistCode As String             '�Һŵ���
    Dim strRegisterOffice As String         '�������
    Dim strRegisterDoctor As String         'ҽ��
    Dim rsTemp As New ADODB.Recordset
    Dim STR����ʱ�� As String
    STR����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If STR����ʱ�� > "2007-04-01" Then
       Exit Function
    End If
       
    On Error GoTo errHand
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    
      
    strIdentify = frmIdentify��Ϫũҽ.GetPatient(bytType, lng����ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0 Or bytType = 3) Then Exit Function
    
    '��������Ǽ�
    gComInfo_��Ϫũҽ.���㴮 = ""
    If bytType = 0 Then
        '����Ƿ��Ѿ�������
        'Balx    varchar(2)  ��ҽ����(=1 ���ﱸ�� =2 תԺ���� =3 ���ⲡ�ֱ���)
        'Kzhm    varchar(30) ��֤����
        '�������ݰ�����
        'Returncode  Long    �������0��ʾ�ɹ�
        'Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
        'Spbz    varchar(1)  ������־=0δ���� =1 ����ͨ�� =2����δͨ��
       ' StrInput = "Balx=3" & mstrSplit & "Kzhm=" & gComInfo_��Ϫũҽ.���˱��
        
       ' Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_AllowQuery, StrInput)
        'If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
        '����Ƿ�ͨ������
        'arrReturn = Split(gstrOutput_��Ϫũҽ, mstrSplit)
        'intVerify = Val(Split(arrReturn(2), "=")(1))
        'Select Case intVerify
        'Case 0
         '   MsgBox "ũҽ�컹δ�������������������ҵ��", vbInformation, gstrSysName
            'Exit Function
        'Case 2
         '   MsgBox "ũҽ��û��ͨ���������������������ҵ��", vbInformation, gstrSysName
            'Exit Function
        'End Select
        
        '��Σ�����ҽ�ƺ��멦����ҽ�Ʋ�����ҽԺ����ĹҺź��멦�����ҽ�����ҽԺ����Ŀ��ҩ������ҽ����" & _
        ҽԺ����ϩ�ҽԺ����Ǽǵ����ک�����֢����������Ļ������멦��������Ļ������Ʃ����쵥λ��������
        'ȡ����ҺŵĿ�����ҽ��
        gstrSQL = " Select B.���� AS �Һſ���,ִ���� AS ҽ�� " & _
                  " From ������ü�¼ A,���ű� B " & _
                  " Where A.��¼����=4 And A.��¼״̬=1 And A.����ID=" & lng����ID & _
                  " And A.ִ�в���ID=B.ID And A.�Ǽ�ʱ�� Between to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 00:00:00','yyyy-MM-dd hh24:mi:ss')" & _
                  " And to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss') And Rownum<2"
      '  Call OpenRecordset(rsTemp, "ȡ����ҺŵĿ�����ҽ��")
      '  If rsTemp.RecordCount = 0 Then
      '      MsgBox "����û����Ч�ĹҺż�¼,�޷������������Ǽǣ�", vbInformation, gstrSysName
       '     Exit Function
      '  End If
        'strRegisterOffice = rsTemp!�Һſ���
         'strRegisterDoctor = rsTemp!ҽ��
        strRegisterOffice = "001"
        strRegisterDoctor = "001"
        'ȡ��������
        gstrSQL = "Select ���� From ��������Ŀ¼ Where ID=(Select nvl(����ID,0) From �����ʻ� Where ����=[1] And ����ID=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", TYPE_��Ϫũҽ, lng����ID)
        If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!����
        
        '��ȡ�Һŵ��ţ�ʮλ��Ψһ��ʶ
        strRegistCode = CStr(zlDatabase.GetNextID("���ű�"))
         
        gComInfo_��Ϫũҽ.������ˮ�� = strRegistCode
        '�������Ǽ����
'        Mzdjh   varchar(20) ����ǼǺţ���ҽԺ���ݿ��е�Ψһ����
'        Kzhm    varchar(30) ��֤����
'        Jzks    varchar(20) �������
'        Ysxm    varchar(10) ҽ������
'        Bzdm    varchar(20) ���ִ���
'        Mzrq    Datetime    ��������(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
'        Czy     Varchar(10) ����Ա����
      
        StrInput = "Mzdjh=" & strRegistCode & mstrSplit & "Kzhm=" & gComInfo_��Ϫũҽ.���˱�� & mstrSplit & _
            "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
            "Bzdm=" & strDiseaseCode & mstrSplit & "Mzrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
         
        If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
           Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_OutRegist, StrInput)
        
           If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
        End If
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ϫũҽ & ",'ҵ������','''" & gComInfo_��Ϫũҽ.ҵ������ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    End If
    
    If bytType = 1 Then
        '���±����ʻ������Ϣ��ҵ�����ͣ�
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ϫũҽ & ",'ҵ������','''" & gComInfo_��Ϫũҽ.ҵ������ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    End If
    
    '���ز�����Ϣ��
    ��ݱ�ʶ_��Ϫũҽ = strIdentify
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ʼ��_��Ϫũҽ(Optional ByVal blnTest As Boolean = False) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim lngPort As Long
    Dim strIP As String
    Dim rsTemp As New ADODB.Recordset
    Dim cnTest As New ADODB.Connection

    On Error Resume Next
    
    If mblnInit = False Then
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_��Ϫũҽ)
        gComInfo_��Ϫũҽ.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        
        'ȡ���ղ���
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", TYPE_��Ϫũҽ)
        Do While Not rsTemp.EOF
            Select Case rsTemp!������
            Case "IP��ַ"
                strIP = Nvl(rsTemp!����ֵ, "127.0.0.1")
            Case "�˿ں�"
                lngPort = Nvl(rsTemp!����ֵ, 8801)
            End Select
            rsTemp.MoveNext
        Loop
        
        '�����Ƿ����ӵ�ͨ
'        Call gobjCXXNY.CXNY_SetRemoteServerAddr(lngPort, strIP)
'yjj1row
        Call CXNY_SetRemoteServerAddr(lngPort, strIP)
        mblnInit = True
    End If
    
    ҽ����ʼ��_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_��Ϫũҽ() As Boolean
    On Error Resume Next
    
    ҽ����ֹ_��Ϫũҽ = True
End Function

Public Function �����������_��Ϫũҽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim StrInput As String, strinput1 As String
    Dim lng����ID As Long, ������ As Long
    Dim str�������� As String, str������ As String, strҽ������ As String, str��Ŀ���� As String, str��� As String, str��λ As String
    Dim dbl�ʻ�֧�� As Double, dbl�ֽ� As Double, dblͳ����� As Double
    Dim fpze As Double, ylzl As Double, blzf As Double
    Dim qfbz As Double, sbje As Double, Zhzf As Double
    Dim Grzf As Double, Zhye As Double, dnbxlj As Double
    Dim cfdkbje As Double
    Dim jsdh As Double
    Dim tbkbje As Double
    Dim tbsbje As Double
    Dim strRegisterOffice  As String
    Dim strRegisterDoctor As String
    Dim strDiseaseCode As String
    Dim strRegistCode As String
    Dim lngrowcount As String
    
    
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    
    '������ϸ������Σ�������û�иú������Է���Ӧ����һ���Ƿ��ṩ��
   ' If gComInfo_��Ϫũҽ.���㴮 <> "" Then
        
    '    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_DelAllDetail1, gComInfo_��Ϫũҽ.���㴮��)
     '   If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    'End If
    
    lng����ID = rs��ϸ!����ID
    
    ''''''''''''''''''
    
      gComInfo_��Ϫũҽ.���㴮 = ""
    
        
        '��Σ�����ҽ�ƺ��멦����ҽ�Ʋ�����ҽԺ����ĹҺź��멦�����ҽ�����ҽԺ����Ŀ��ҩ������ҽ����" & _
        ҽԺ����ϩ�ҽԺ����Ǽǵ����ک�����֢����������Ļ������멦��������Ļ������Ʃ����쵥λ��������
       ' ȡ����ҺŵĿ�����ҽ��
       ' strRegisterOffice = rs��ϸ!��������ID
          strRegisterOffice = "00001"
         strRegisterDoctor = rs��ϸ!������
         
       'gstrSQL = " Select ���� AS �Һſ��� from ���ű� " & _
      '            " Where ID=" & strRegisterOffice
      '           WriteInfo (gstrSQL)
      '  Call OpenRecordset(rsTemp, "ȡ����ҺŵĿ���")
  '    If rsTemp.RecordCount = 0 Then
   '         MsgBox "����û����Ч�ĹҺż�¼,�޷������������Ǽǣ�", vbInformation, gstrSysName
   '    Exit Function
    '   End If
     '   strRegisterOffice = rsTemp!�Һſ���
      '
        'ȡ��������
        gstrSQL = "Select ���� From ��������Ŀ¼ Where ID=(Select nvl(����ID,0) From �����ʻ� Where ����=[1] And ����ID=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", TYPE_��Ϫũҽ, lng����ID)
        If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!����
        gComInfo_��Ϫũҽ.�������� = strDiseaseCode
        '��ȡ�Һŵ��ţ�ʮλ��Ψһ��ʶ
        strRegistCode = CStr(zlDatabase.GetNextID("���ű�"))
         
        gComInfo_��Ϫũҽ.������ˮ�� = strRegistCode
        '�������Ǽ����
'        Mzdjh   varchar(20) ����ǼǺţ���ҽԺ���ݿ��е�Ψһ����
'        Kzhm    varchar(30) ��֤����
'        Jzks    varchar(20) �������
'        Ysxm    varchar(10) ҽ������
'        Bzdm    varchar(20) ���ִ���
'        Mzrq    Datetime    ��������(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
'        Czy     Varchar(10) ����Ա����
        strinput1 = "Mzdjh=" & strRegistCode & mstrSplit & "Kzhm=" & gComInfo_��Ϫũҽ.���˱�� & mstrSplit & _
            "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
            "Bzdm=" & strDiseaseCode & mstrSplit & "Mzrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
        
        If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
           Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_OutRegist, strinput1)
           If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
        End If
     
    ''''''''''''''''''''
       str�������� = Format(zlDatabase.Currentdate, mstrDateFormat)
        
    '�õ����ν�����ܷ���
    With rs��ϸ
        '������ܶ�
        gComInfo_��Ϫũҽ.�ܷ��� = 0
        Do While Not .EOF
            gComInfo_��Ϫũҽ.�ܷ��� = gComInfo_��Ϫũҽ.�ܷ��� + !ʵ�ս��
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
          lngrowcount = 0
        Do While Not .EOF
            '��ȡ�շ�ϸĿ�������Ϣ
            If Nvl(!ʵ�ս��, 0) <> 0 Then
            lngrowcount = lngrowcount + 1
             gstrSQL = " Select A.��� AS �շ����,A.����,A.���,A.���㵥λ AS ��λ,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                      " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=[1] And A.ID=[2]"
            Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", TYPE_��Ϫũҽ, CLng(!�շ�ϸĿID))
           
            
            str��Ŀ���� = Nvl(rsItem!����, "��")
             If IsNull(rsItem!��Ŀ����) = True Or rsItem!��Ŀ���� = "" Then
                MsgBox "ҽԺ��Ŀ" & str��Ŀ���� & "û�ж��룬���ȶ��룡"
                Exit Function
            End If
            strҽ������ = Nvl(rsItem!��Ŀ����)
            str��� = Nvl(rsItem!���, "��")
            str��λ = Nvl(rsItem!��λ, "��")
            If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
            str��� = Nvl(str���, "��")
            '������ϸ�ϴ����
        '    Jylx    varchar(1)  ��ҽ����(=0 ���=1 סԺ)
        '    Jyhm    varchar(20) ��ҽ����(��Ժ�ǼǺŻ�����ǼǺ�)
        '    Recordcount numeric(10) ��¼����(��=100)
        '    xmdm[i] varchar(10) ��Ŀ����(��ũҽ��Ŀ����)
        '    zxxmmc[i]   varchar(80) ������Ŀ����(ҽԺ��Ŀ����)
        '    xmdw[i] varchar(10) ��Ŀ��λ
        '    xmdj[i] numeric(10,4)   ��Ŀ����
        '    xmsl[i] numeric(10,4)   ��Ŀ����
        '    xmje[i] numeric(10,4)   ��Ŀ���
        '    xmgg[i] varchar(50) ��Ŀ���
        '    yzrq[i] Datetime    ҽ�����ڻ���������(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
            
            StrInput = StrInput & mstrSplit & "xmdm[" & lngrowcount & "]=" & strҽ������ & mstrSplit & "zxxmmc[" & lngrowcount & "]=" & ToVarchar(str��Ŀ����, 80) & mstrSplit & _
                "xmdw[" & lngrowcount & "]=" & ToVarchar(str��λ, 10) & mstrSplit & "xmdj[" & lngrowcount & "]=" & Format(!����, mstrPriceFormat) & mstrSplit & _
                "xmsl[" & lngrowcount & "]=" & Format(!����, mstrAmountFormat) & mstrSplit & "xmje[" & lngrowcount & "]=" & Format(!ʵ�ս��, mstrMoneyFormat) & mstrSplit & _
                "xmgg[" & lngrowcount & "]=" & ToVarchar(str���, 50) & mstrSplit & "yzrq[" & lngrowcount & "]=" & str��������
             End If
     
            .MoveNext
            
          
        Loop
         If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
            StrInput = "Jylx=0" & mstrSplit & "Jyhm=" & gComInfo_��Ϫũҽ.������ˮ�� & mstrSplit & "Recordcount=" & lngrowcount & StrInput
             Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
             If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
         End If
    End With
    'gComInfo_��Ϫũҽ.���㴮�� = strinput1
    'Ԥ��������
'    Jsbj    Varchar(1)  ������(=0 Ԥ�� =1 ����)
'    Mzdjh   varchar(20) ����ǼǺ�
'    Kzhm    varchar(30) ��֤����
'    Jsrq    Datetime    ��������(��������) (�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)
'    Recordcount Long    ��ϸ��¼��
'    Fpze    numeric(10,2)   ��Ʊ�ܶ�
'    Czy Varchar(10) ����Ա����
    StrInput = "Jsbj=0" & mstrSplit & "Mzdjh=" & gComInfo_��Ϫũҽ.������ˮ�� & mstrSplit & "Kzhm=" & gComInfo_��Ϫũҽ.���˱�� & mstrSplit & _
        "Jsrq=" & str�������� & mstrSplit & "Recordcount=" & lngrowcount & mstrSplit & "Fpze=" & Format(gComInfo_��Ϫũҽ.�ܷ���, "#0.00") & mstrSplit & "Czy=" & gstrUserName
    
    gComInfo_��Ϫũҽ.���㴮 = StrInput
    gComInfo_��Ϫũҽ.���㴮�� = "Mzdjh=" & gComInfo_��Ϫũҽ.������ˮ�� & mstrSplit & "Kzhm=" & gComInfo_��Ϫũҽ.���˱��
    
    If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
       Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_OutBalance, StrInput)
       If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    End If
    '���Σ�
'    Returncode  Long    �������0��ʾ�ɹ�
'    Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
'    Fpze    numeric(10,2)   ��Ʊ�ܶ�
'    Ylzl    numeric(10,2)   ��������
'    Blzf    numeric(10,2)   �����Է�
'    Qfbz    numeric(10,2)   �𸶱�׼
'    Sbje    numeric(10,2)   ʵ�����
'    Zhzf    numeric(10,2)   �ʻ�֧��
'    Grzf    numeric(10,2)   �����Ը�
'    Zhye    numeric(10,2)   �ʻ����
'    Dnbxlj  Numeric(10,2)   ����ͳ�ﱨ���ۼ�
'    Cfdkbje numeric(10,2)   ���ⶥ��Ч���(���ڱ�����)
    If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
        gComInfo_��Ϫũҽ.�ܷ���_���� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(2), "=")(1), "#0.00"))
        If Format(gComInfo_��Ϫũҽ.�ܷ���_����, "#0.00") <> Format(gComInfo_��Ϫũҽ.�ܷ���, "#0.00") Then
            Err.Raise 9000, gstrSysName, "ҽԺ���ܷ�����ҽ�����ĵ��ܷ��ò�һ�£�" & vbCrLf & _
            "ҽԺ��" & Format(gComInfo_��Ϫũҽ.�ܷ���, "#0.00") & Space(10) & "ҽ�����ģ�" & Format(gComInfo_��Ϫũҽ.�ܷ���_����, "#0.00"), vbInformation, gstrSysName
        End If
    
        dblͳ����� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(6), "=")(1), "#0.00"))
        dbl�ʻ�֧�� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(7), "=")(1), "#0.00"))
        fpze = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(2), "=")(1), "#0.00"))
        ylzl = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(3), "=")(1), "#0.00"))
        blzf = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(4), "=")(1), "#0.00"))
        qfbz = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(5), "=")(1), "#0.00"))
        Grzf = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(8), "=")(1), "#0.00"))
        Zhye = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(9), "=")(1), "#0.00"))
        dnbxlj = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(10), "=")(1), "#0.00"))
        cfdkbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(11), "=")(1), "#0.00"))
        jsdh = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(12), "=")(1), "#0.00"))
        tbkbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(13), "=")(1), "#0.00"))
        tbsbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(14), "=")(1), "#0.00"))
        dbl�ֽ� = gComInfo_��Ϫũҽ.�ܷ��� - dblͳ����� - dbl�ʻ�֧��
    Else
        dbl�ʻ�֧�� = 0
        dblͳ����� = 0
    End If
    str���㷽ʽ = "�����ʻ�;" & dbl�ʻ�֧�� & ";0|ͳ�����;" & dblͳ����� & ";0"
      
    �����������_��Ϫũҽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �������_��Ϫũҽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    Dim lng����ID As Long
    Dim StrInput As String, str���㵥�� As String
    
    Dim dbl���� As Double, dblͳ����� As Double, dbl�ֽ� As Double
        Dim fpze As Double, ylzl As Double, blzf As Double
    Dim qfbz As Double, sbje As Double, Zhzf As Double
    Dim Grzf As Double, Zhye As Double, dnbxlj As Double
    Dim cfdkbje As Double
    Dim jsdh As Double
    Dim tbkbje As Double
    Dim tbsbje As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    StrInput = gComInfo_��Ϫũҽ.���㴮
    StrInput = Replace(StrInput, "Jsbj=0", "Jsbj=1")
     If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
        Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_OutBalance, StrInput)
        If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
     End If
    '���Σ�
'    Returncode  Long    �������0��ʾ�ɹ�
'    Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
'    Fpze    numeric(10,2)   ��Ʊ�ܶ�
'    Ylzl    numeric(10,2)   ��������
'    Blzf    numeric(10,2)   �����Է�
'    Qfbz    numeric(10,2)   �𸶱�׼
'    Sbje    numeric(10,2)   ʵ�����
'    Zhzf    numeric(10,2)   �ʻ�֧��
'    Grzf    numeric(10,2)   �����Ը�
'    Zhye    numeric(10,2)   �ʻ����
'    Dnbxlj  Numeric(10,2)   ����ͳ�ﱨ���ۼ�
'    Cfdkbje numeric(10,2)   ���ⶥ��Ч���(���ڱ�����)
'   Jsdh    numeric(15) ���㵥��
'   Tbkbje  numeric(10,2)   �ز��ɱ����
'   Tbsbje  numeric(10,2)   �ز�ʵ�����

      'ȡ����ID
    gstrSQL = "Select ����ID From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵�ID", lng����ID)
    lng����ID = rsTemp!����ID
    If gComInfo_��Ϫũҽ.�������� <> "�º󲹱�" Then
        dbl���� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(5), "=")(1), "#0.00"))
        dblͳ����� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(6), "=")(1), "#0.00"))
        cur�����ʻ� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(7), "=")(1), "#0.00"))
        dbl�ֽ� = gComInfo_��Ϫũҽ.�ܷ���_���� - dblͳ����� - cur�����ʻ�
        str���㵥�� = str(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(12), "=")(1), "#0"))
         fpze = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(2), "=")(1), "#0.00"))
        ylzl = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(3), "=")(1), "#0.00"))
        blzf = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(4), "=")(1), "#0.00"))
        qfbz = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(5), "=")(1), "#0.00"))
        Grzf = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(8), "=")(1), "#0.00"))
        Zhye = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(9), "=")(1), "#0.00"))
        dnbxlj = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(10), "=")(1), "#0.00"))
        cfdkbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(11), "=")(1), "#0.00"))
        jsdh = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(12), "=")(1), "#0.00"))
        tbkbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(13), "=")(1), "#0.00"))
        tbsbje = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(14), "=")(1), "#0.00"))
 
    
    Else
       dbl���� = 0
       dblͳ����� = 0
       dbl�ֽ� = 0
       str���㵥�� = "�º󲹱�"
        gstrSQL = "insert into ���˷��ü�¼_�º󲹱���ϵ Select * From ������ü�¼ Where ����ID=" & lng����ID
        gcnOracle.Execute gstrSQL
        gstrSQL = "delete from ���˷��ü�¼_�º󲹱���ϵ  Where �Ƿ��ϴ�=1 and ����ID=" & lng����ID
       gcnOracle.Execute gstrSQL
        
       gstrSQL = "update ���˷��ü�¼_�º󲹱���ϵ set ִ����='" & gComInfo_��Ϫũҽ.�������� & "' Where ����ID=" & lng����ID
       gcnOracle.Execute gstrSQL
        
    End If
         
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ϫũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
         dnbxlj & "," & "NULL" & "," & qfbz & "," & 0 & "," & 0 & "," & _
        gComInfo_��Ϫũҽ.�ܷ��� & "," & blzf & "," & ylzl & "," & fpze - blzf - ylzl & "," & dblͳ����� & "," & tbkbje & " ," & tbsbje & "," & _
        cur�����ʻ� & ",'" & gComInfo_��Ϫũҽ.������ˮ�� & "',null,null,'" & str���㵥�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    gComInfo_��Ϫũҽ.���㴮 = ""
    �������_��Ϫũҽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_��Ϫũҽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str��֤���� As String
    Dim StrInput As String
    Dim str���㵥��
    Dim rsTemp As New ADODB.Recordset, rsTemp1 As New ADODB.Recordset
    On Error GoTo errHand
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    'ֻ�������һ��
    'ȡ������¼�Ľ���ID�����ݺ�
    
    'ȡ��֤����
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����֤����", TYPE_��Ϫũҽ, lng����ID)
    str��֤���� = Nvl(rsTemp!����)
    
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=1 And ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ˮ��", lng����ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ�����¼���޷�����������������", vbInformation, gstrSysName
        Exit Function
    End If
    gComInfo_��Ϫũҽ.������ˮ�� = Nvl(rsTemp!֧��˳���)
    str���㵥�� = Nvl(rsTemp!��ע)
    
    '���ý������
'    Jylx    varchar(1)  ��ҽ����(=0 �������� = 1 ȡ����Ժ)
'    Jyhm    varchar(20) ��ҽ����
'    Kzhm    varchar(30) ��֤����
'    Tfrq    Datetime    ��������(�˷�����) (�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)
'    Czy     Varchar(10) ����Ա����
   If str���㵥�� <> "�º󲹱�" Then
        StrInput = "Jylx=0" & mstrSplit & "Jyhm=" & gComInfo_��Ϫũҽ.������ˮ�� & mstrSplit & _
        "Kzhm=" & str��֤���� & mstrSplit & "Tfrq=" & Format(zlDatabase.Currentdate(), mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
        Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_BalanceCancel, StrInput)
        If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    Else '�º󲹱������Ҫ�˷�Ҫ���ж��Ƿ��Ѿ�����������Ѿ���������Ҫ���ڲ��������˷ѣ�Ȼ���������˷�
       gstrSQL = "Select * From ���˷��ü�¼_�º󲹱���ϵ Where �Ƿ��ϴ�='1' and ����ID=[1]"
       Set rsTemp1 = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�º󲹱���¼��״̬", lng����ID)
       If rsTemp1.RecordCount > 0 Then '˵���Ѿ��ᱨ���������˷�
           MsgBox "�õ������º�ᱨ���ݣ������Ѿ����㣬�����˷ѣ�"
           Exit Function
       End If
       gstrSQL = "update ���˷��ü�¼_�º󲹱���ϵ set �Ƿ��ϴ�='8' Where ����ID=" & lng����ID
       gcnOracle.Execute gstrSQL
    End If
    '���汾�ν������
       
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ϫũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",0," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & rsTemp!֧��˳��� & "',null,null,'" & rsTemp!��ע & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
    
    ����������_��Ϫũҽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_��Ϫũҽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim StrInput As String
    Dim strKey As String                    '��Ժ�ǼǺ�(����=21)
    Dim strRegister As String               '��������
    Dim strCardNO As String, lngDisease As Long '���˵�ҽ�ƿ��ż�����ID
    Dim strRegistCode As String             '�Һŵ���
    Dim strInHospitalDate As String         '��Ժ����
    Dim strRegisterOffice As String         '�������
    Dim strDiseaseCode As String            '���ִ���
    Dim strDiagnose As String               '��Ժ���
    Dim strRegisterDoctor As String         'ҽ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strKey = GetKey(lng����ID, lng��ҳID)
    'ȡ���˵�ҽ�ƿ���
    gstrSQL = "Select ����,Nvl(����ID,0) ����ID,Nvl(ҵ������,'21') AS �������� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�ҽ�ƿ���", TYPE_��Ϫũҽ, lng����ID)
    strCardNO = rsTemp!����
    lngDisease = rsTemp!����ID
    strRegister = Val(rsTemp!��������) - 21
    
    'ȡ������ҽ��
    gstrSQL = " Select A.��Ժ����,B.���� ����,A.����ҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ҽ��", lng����ID, lng��ҳID)
    strInHospitalDate = Format(rsTemp!��Ժ����, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!ҽ��)
    strRegisterOffice = Nvl(rsTemp!����)
    'ȡ���ִ���
    gstrSQL = "Select ���� From ��������Ŀ¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ִ���", lngDisease)
    If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!����
    'ȡ��Ժ���
    strDiagnose = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, False)
    
    '��Σ�
'    Rydjh   varchar(20) ��Ժ�ǼǺţ���ҽԺ���ݿ��е�Ψһ����
'    Kzhm    varchar(30) ��֤����
'    Jzks    varchar(20) �������
'    Jslx    Varchar(1)  ��������(=0 ��ͨ =1 ��ͨ�¹� =2 �󲡾��� =3 �Ѳ� =4 ����)
'    Ysxm    varchar(10) ҽ������
'    Bzdm    Varchar(20) ���ִ���
'    Ryzdsm  varchar(254)    ��Ժ���˵��
'    Ryrq    Datetime    ��Ժ����(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
'    Czy Varchar(10) ����Ա����
    '���ز���:
'    Returncode  Long    �������0��ʾ�ɹ�
'    Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & _
        "Jzks=" & strRegisterOffice & mstrSplit & "Jslx=" & strRegister & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
        "Bzdm=" & strDiseaseCode & mstrSplit & "Ryzdsm=" & strDiagnose & mstrSplit & _
        "Ryrq=" & strInHospitalDate & mstrSplit & "Czy=" & UserInfo.����
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_InRegist, StrInput)
    If Not ���ýӿ�_��Ϫũҽ() Then Exit Function

    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ϫũҽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")

    ��Ժ�Ǽ�_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��Ϫũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim StrInput As String
    Dim strKey As String, strCardNO As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�����ҳIDС��һλ����ǰ���һ����
    strKey = GetKey(lng����ID, lng��ҳID)
    
    'ȡ���˵�ҽ�ƿ���
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�ҽ�ƿ���", TYPE_��Ϫũҽ, lng����ID)
    strCardNO = rsTemp!����

    '��Σ�
'    Rydjh   varchar(20) ��Ժ�ǼǺţ���ҽԺ���ݿ��е�Ψһ����
'    Kzhm    varchar(30) ��֤����
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_InRegistCancel, StrInput)
    If Not ���ýӿ�_��Ϫũҽ Then Exit Function

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ϫũҽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    gstrSQL = "zl_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    ��Ժ�Ǽǳ���_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_��Ϫũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    
    Call ������Ժ��Ϣ_��Ϫũҽ(lng����ID, lng��ҳID, True)
    '�������Ա�޸ľ������ͣ�ͬʱ����
    Call frm���������޸�.ShowME(lng����ID)
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ϫũҽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��Ϫũҽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ϫũҽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_��Ϫũҽ(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '����: ��ȡ�α����˸����ʻ����
    '����: strSelfNO-���˸��˱��
    '����: ���ظ����ʻ����Ľ��
    '�����������ؼ�ͥ�ʻ���סԺ���ظ����ʻ����
    gstrSQL = "Select Nvl(�ʻ����,0) AS �����ʻ� From �����ʻ� Where ҽ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", strSelfNo)
    �������_��Ϫũҽ = rsTemp!�����ʻ�
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����ϴ�_��Ϫũҽ(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal strNO As String) As Boolean
    Dim intCOUNT As Integer
    Dim lng��ҳID As Long, lng����ID As Long
    Dim StrInput As String
    Dim blnInsure As Boolean, blnTrans As Boolean
    Dim str��Ŀ���� As String, strҽ������ As String, str��� As String, str��λ As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    '�ϴ�������ϸ��������ɺ��ϴ�����һ����������������¼��δ������Ŀ��Ȼ��һһ���ϴ���ǣ�
    '�򿪱��δ��ϴ��Ĵ�����ϸ
    gstrSQL = " Select A.ID,A.��¼����,A.��¼״̬,A.NO,A.���,A.�շ����,A.����ID,A.��ҳID,A.�շ�ϸĿID,A.�Ǽ�ʱ��,A.ʵ�ս��," & _
              " Nvl(A.����,1)*A.���� AS ����,A.ʵ�ս��/(Nvl(A.����,1)*A.����) AS �۸�" & _
              " From סԺ���ü�¼ A,�����ʻ� B" & _
              " Where A.��¼����=" & int���� & " ANd A.��¼״̬=" & int״̬ & " And A.NO='" & strNO & "' And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.ʵ�ս��,0)<>0 " & _
              " And A.����ID=B.����ID And B.����=" & TYPE_��Ϫũҽ & _
              " Order by A.����ID"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���δ��ϴ��Ĵ�����ϸ")

    '������ɺ��ϴ�
    gcnOracle.BeginTrans
    blnTrans = True
    With rsDetail
        '�ϴ�����
        lng����ID = 0
        Do While Not .EOF
            If lng����ID <> !����ID Then
                If lng����ID <> 0 Then
                    '˵��������ϸ����װ���ˣ�׼���ϴ�
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(!����ID, !��ҳID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
                    If ���ýӿ�_��Ϫũҽ() Then
                        gcnOracle.CommitTrans
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
                
                intCOUNT = 0
                StrInput = ""
                lng����ID = !����ID
                lng��ҳID = !��ҳID
                blnInsure = IsYBPatient(lng����ID)
            End If

            If blnInsure Then
                '��ȡ�շ�ϸĿ�������Ϣ
                gstrSQL = " Select A.��� AS �շ����,A.����,A.���,A.���㵥λ AS ��λ,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                          " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=[1] And A.ID=[2]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", TYPE_��Ϫũҽ, CLng(!�շ�ϸĿID))
                str��Ŀ���� = Nvl(rsItem!����)
                strҽ������ = Nvl(rsItem!��Ŀ����)
                strҽ������ = getС����ҩ����(lng����ID, !�շ�ϸĿID, strҽ������)
                str��� = Nvl(rsItem!���)
                str��λ = Nvl(rsItem!��λ)
                If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
        
                '������ϸ�ϴ����
            '    Jylx    varchar(1)  ��ҽ����(=0 ���=1 סԺ)
            '    Jyhm    varchar(20) ��ҽ����(��Ժ�ǼǺŻ�����ǼǺ�)
            '    Recordcount numeric(10) ��¼����(��=100)
            '    xmdm[i] varchar(10) ��Ŀ����(��ũҽ��Ŀ����)
            '    zxxmmc[i]   varchar(80) ������Ŀ����(ҽԺ��Ŀ����)
            '    xmdw[i] varchar(10) ��Ŀ��λ
            '    xmdj[i] numeric(10,4)   ��Ŀ����
            '    xmsl[i] numeric(10,4)   ��Ŀ����
            '    xmje[i] numeric(10,4)   ��Ŀ���
            '    xmgg[i] varchar(50) ��Ŀ���
            '    yzrq[i] Datetime    ҽ�����ڻ���������(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
                
                intCOUNT = intCOUNT + 1
                StrInput = StrInput & mstrSplit & "xmdm[" & intCOUNT & "]=" & strҽ������ & mstrSplit & "zxxmmc[" & intCOUNT & "]=" & ToVarchar(str��Ŀ����, 80) & mstrSplit & _
                    "xmdw[" & intCOUNT & "]=" & ToVarchar(str��λ, 10) & mstrSplit & "xmdj[" & intCOUNT & "]=" & Format(!�۸�, mstrPriceFormat) & mstrSplit & _
                    "xmsl[" & intCOUNT & "]=" & Format(!����, mstrAmountFormat) & mstrSplit & "xmje[" & intCOUNT & "]=" & Format(!ʵ�ս��, mstrMoneyFormat) & mstrSplit & _
                    "xmgg[" & intCOUNT & "]=" & ToVarchar(str���, 50) & mstrSplit & "yzrq[" & intCOUNT & "]=" & Format(!�Ǽ�ʱ��, mstrDateFormat)
                
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                If intCOUNT = 20 Then
                    '˵��������ϸ����װ���ˣ�׼���ϴ�
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(lng����ID, lng��ҳID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
                    If ���ýӿ�_��Ϫũҽ() Then
                        gcnOracle.CommitTrans
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                    intCOUNT = 0
                    StrInput = ""
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If intCOUNT <> 0 Then
        '˵��������ϸ����װ���ˣ�׼���ϴ�
        StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(lng����ID, lng��ҳID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
        Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
        If ���ýӿ�_��Ϫũҽ() Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    blnTrans = False

    �����ϴ�_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function סԺ�������_��Ϫũҽ(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim StrInput As String
    Dim strKey As String
    Dim blnTrans As Boolean
    Dim strCardNO As String, str��Ժ���� As String, str��Ժ��� As String
    Dim intBalance As Integer, intRecords As Integer, intCOUNT As Integer
    Dim lng��ҳID As Long, dbl�е����� As Double
    Dim dbl�ʻ�֧�� As Double, dbl�ֽ� As Double, dblҽ������ As Double, dbl���� As Double
    Dim str��Ŀ���� As String, strҽ������ As String, str��� As String, str��λ As String
    Dim rsItem As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim rs��ϸ1 As New ADODB.Recordset
    On Error GoTo errHand
    
    'ȡ��������
    gstrSQL = "Select ����,ҵ������,Nvl(����֤��,0) �е����� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", TYPE_��Ϫũҽ, lng����ID)
    intBalance = Val(Right(rsItem!ҵ������, 1)) - 1
    dbl�е����� = Val(rsItem!�е�����)
    strCardNO = rsItem!����

    'ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID")
    lng��ҳID = rsItem!��ҳID
    strKey = GetKey(lng����ID, lng��ҳID)
    
    'ȡ��Ժ����
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ����", lng����ID, lng��ҳID)
    str��Ժ���� = Format(rsItem!��Ժ����, mstrDateFormat)
    If str��Ժ���� = "" Then str��Ժ���� = Format(zlDatabase.Currentdate, mstrDateFormat)
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, False)

    '��ȡ���η�����ϸ
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ����,A.������ AS ҽ��,A.�Ǽ�ʱ�� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 And A.����Ա���� is not null AND Nvl(A.ʵ�ս��,0)<>0 " & _
              "        And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= [3]" & _
              "  Order by A.����ID,A.����ʱ��"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", lng����ID, lng��ҳID, TYPE_��Ϫũҽ)

 '��ȡ���η�����ϸ1
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ����,A.������ AS ҽ��,A.�Ǽ�ʱ�� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 And A.����Ա���� is not null AND Nvl(A.ʵ�ս��,0)<>0 " & _
              "        And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= [3]" & _
              "  Order by A.����ID,A.����ʱ��"
    Set rs��ϸ1 = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", lng����ID, lng��ҳID, TYPE_��Ϫũҽ)

    With rs��ϸ1
        '������ܶ�
        gComInfo_��Ϫũҽ.�ܷ��� = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!���, 0) <> 0 Then
                intRecords = intRecords + 1
                gComInfo_��Ϫũҽ.�ܷ��� = gComInfo_��Ϫũҽ.�ܷ��� + !���
            End If
            .MoveNext
        Loop
        
    End With
  
    gcnOracle.BeginTrans
    blnTrans = True
    With rs��ϸ
        Do While Not .EOF
            If Nvl(!���, 0) <> 0 Then
               '��ȡ�շ�ϸĿ�������Ϣ
                gstrSQL = " Select A.��� AS �շ����,A.����,A.���,A.���㵥λ AS ��λ,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B" & _
                          " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=[1] And A.ID=[2]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", TYPE_��Ϫũҽ, CLng(!�շ�ϸĿID))
                str��Ŀ���� = Nvl(rsItem!����)
                strҽ������ = Nvl(rsItem!��Ŀ����)
                strҽ������ = getС����ҩ����(lng����ID, !�շ�ϸĿID, strҽ������)
                str��� = Nvl(rsItem!���)
                str��λ = Nvl(rsItem!��λ)
                If InStr(1, str���, "|") <> 0 Then str��� = Mid(str���, 1, InStr(1, str���, "|") - 1)
        
                '������ϸ�ϴ����
            '    Jylx    varchar(1)  ��ҽ����(=0 ���=1 סԺ)
            '    Jyhm    varchar(20) ��ҽ����(��Ժ�ǼǺŻ�����ǼǺ�)
            '    Recordcount numeric(10) ��¼����(��=100)
            '    xmdm[i] varchar(10) ��Ŀ����(��ũҽ��Ŀ����)
            '    zxxmmc[i]   varchar(80) ������Ŀ����(ҽԺ��Ŀ����)
            '    xmdw[i] varchar(10) ��Ŀ��λ
            '    xmdj[i] numeric(10,4)   ��Ŀ����
            '    xmsl[i] numeric(10,4)   ��Ŀ����
            '    xmje[i] numeric(10,4)   ��Ŀ���
            '    xmgg[i] varchar(50) ��Ŀ���
            '    yzrq[i] Datetime    ҽ�����ڻ���������(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
                
                intCOUNT = intCOUNT + 1
                StrInput = StrInput & mstrSplit & "xmdm[" & intCOUNT & "]=" & strҽ������ & mstrSplit & "zxxmmc[" & intCOUNT & "]=" & ToVarchar(str��Ŀ����, 80) & mstrSplit & _
                    "xmdw[" & intCOUNT & "]=" & ToVarchar(str��λ, 10) & mstrSplit & "xmdj[" & intCOUNT & "]=" & Format(!����, mstrPriceFormat) & mstrSplit & _
                    "xmsl[" & intCOUNT & "]=" & Format(!����, mstrAmountFormat) & mstrSplit & "xmje[" & intCOUNT & "]=" & Format(!���, mstrMoneyFormat) & mstrSplit & _
                    "xmgg[" & intCOUNT & "]=" & ToVarchar(str���, 50) & mstrSplit & "yzrq[" & intCOUNT & "]=" & Format(!�Ǽ�ʱ��, mstrDateFormat)
                    
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                If intCOUNT = 20 Then
                    '˵��������ϸ����װ���ˣ�׼���ϴ�
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & strKey & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
                    If Not ���ýӿ�_��Ϫũҽ() Then
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                    gcnOracle.CommitTrans
                    gcnOracle.BeginTrans
                    intCOUNT = 0
                    StrInput = ""
                End If
            End If
            .MoveNext
        Loop
    
        If intCOUNT <> 0 Then
            '˵��������ϸ����װ���ˣ�׼���ϴ�
            StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & strKey & mstrSplit & "Recordcount=" & intCOUNT & StrInput
            Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_UploadDetail, StrInput)
            If Not ���ýӿ�_��Ϫũҽ() Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
        End If
    End With
    blnTrans = False
    
    '��Σ�
    'Jsbj    varchar(1)  ������(=0 Ԥ�� =1 ����)
    'Jslx    Varchar(1)  ��������(=0 ��ͨ =1 ��ͨ�¹� =2 �󲡾��� =3 �Ѳ� =4 ����)
    'Cdbl    Numeric(10,2)   �е�����(��ͨ�¹�)
    'Rydjh   varchar(20) ��Ժ�ǼǺţ���ҽԺ���ݿ�Ψһ
    'Kzhm    varchar(30) ��֤����
    'Cyrq    Datetime    ��Ժ����(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)
    'Cyzdsm  varchar(254)    ��Ժ���˵��
    'Recordcount Long    ��ϸ��¼��
    'Fpze    numeric(10,2)   ��Ʊ�ܶ�
    'Czy Varchar(10) ����Ա����
    '���Σ�
    'Returncode  Long    �������0��ʾ�ɹ�
    'Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
    'Fpze    numeric(10,2)   ��Ʊ�ܶ�
    'Ylzl    numeric(10,2)   ��������
    'Blzf    numeric(10,2)   �����Է�
    'Qfbz    numeric(10,2)   �𸶱�׼
    'Sbje    numeric(10,2)   �������
    'Zhzf    numeric(10,2)   �ʻ�֧��
    'Grzf    numeric(10,2)   �����Ը�
    'Zhye    numeric(10,2)   �ʻ����
    'Dnbxlj  numeric(10,2)   ����ͳ�ﱨ���ۼ�
    'Cfdkbje Numeric(10,2)   ���ⶥ��Ч���
    'Jsdh    Number(15)  ���㵥��
    'Kbje    Number(10,2)    ���οɱ����
    'grzyljkb    Number(10,2)    �����ۼƿɱ����(��������)
    'dc  Char(10)    ��������
    'Ld_bcfd[1]  Number(10,2)    �ֶ�1
    'Ld_bckbje[1]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[1]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[2]  Number(10,2)    �ֶ�2
    'Ld_bckbje[2]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[2]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[3]  Number(10,2)    �ֶ�3
    'Ld_bckbje[3]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[3]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[4]  Number(10,2)    �ֶ�4
    'Ld_bckbje[4]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[4]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[5]  Number(10,2)    �ֶ�5
    'Ld_bckbje[5]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[5]    Number(10,2)    ����ʵ�����

    'ע��(1)��Ʊ�ܶ�����������Ը�
    '�ֽ��� = �����Ը�
    '(2)סԺ����ǰ�����е���ϸ�����Ѿ�����
    '(3)������㽻�׺�סԺ������ͬ���Ĵ������ͷ���ֵ
    StrInput = "Jsbj=0" & mstrSplit & "Jslx=" & intBalance & mstrSplit & "Cdbl=" & IIf(intBalance = 1, dbl�е����� / 100, 0) & mstrSplit & _
        "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & "Cyrq=" & str��Ժ���� & mstrSplit & _
        "Cyzdsm=" & str��Ժ��� & mstrSplit & "recordcount=" & intRecords & mstrSplit & "Fpze=" & gComInfo_��Ϫũҽ.�ܷ��� & mstrSplit & _
        "Czy=" & UserInfo.����
    gComInfo_��Ϫũҽ.���㴮 = StrInput
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_InBalance, StrInput)
    If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    
    dblҽ������ = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(6), "=")(1))
    dbl�ʻ�֧�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(7), "=")(1))
    dbl�ֽ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(8), "=")(1))

    סԺ�������_��Ϫũҽ = "�����ʻ�;" & dbl�ʻ�֧�� & ";0|ҽ������;" & dblҽ������ & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function סԺ����_��Ϫũҽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim StrInput As String
    Dim str������ˮ�� As String
    Dim lng��ҳID As Long
    Dim dbl�ֽ� As Double, dblҽ������ As Double, dbl���� As Double, dbl�ʻ�֧�� As Double
    Dim dbl���οɱ� As Double, dbl�����ۼ� As Double, str�������� As String
    Dim dbl�ֶ�1 As Double, dbl�ֶ�1�ɱ� As Double, dbl�ֶ�1ʵ�� As Double
    Dim dbl�ֶ�2 As Double, dbl�ֶ�2�ɱ� As Double, dbl�ֶ�2ʵ�� As Double
    Dim dbl�ֶ�3 As Double, dbl�ֶ�3�ɱ� As Double, dbl�ֶ�3ʵ�� As Double
    Dim dbl�ֶ�4 As Double, dbl�ֶ�4�ɱ� As Double, dbl�ֶ�4ʵ�� As Double
    Dim dbl�ֶ�5 As Double, dbl�ֶ�5�ɱ� As Double, dbl�ֶ�5ʵ�� As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    '�����ȳ�Ժ�����ܽ��н���
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        Err.Raise 9000, gstrSysName, "�����ȳ�Ժ�����ܽ��н��㣡", vbInformation, gstrSysName
        Exit Function
    End If

    'ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    StrInput = gComInfo_��Ϫũҽ.���㴮
    StrInput = Replace(StrInput, "Jsbj=0", "Jsbj=1")
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_InBalance, StrInput)
    If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    '���Σ�
    'Returncode  Long    �������0��ʾ�ɹ�
    'Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
    'Fpze    numeric(10,2)   ��Ʊ�ܶ�
    'Ylzl    numeric(10,2)   ��������
    'Blzf    numeric(10,2)   �����Է�
    'Qfbz    numeric(10,2)   �𸶱�׼
    'Sbje    numeric(10,2)   �������
    'Zhzf    numeric(10,2)   �ʻ�֧��
    'Grzf    numeric(10,2)   �����Ը�
    'Zhye    numeric(10,2)   �ʻ����
    'Dnbxlj  numeric(10,2)   ����ͳ�ﱨ���ۼ�
    'Cfdkbje Numeric(10,2)   ���ⶥ��Ч���
    'Jsdh    Number(15)  ���㵥��
    'Kbje    Number(10,2)    ���οɱ����
    'grzyljkb    Number(10,2)    �����ۼƿɱ����(��������)
    'dc  Char(10)    ��������
    'Ld_bcfd[1]  Number(10,2)    �ֶ�1
    'Ld_bckbje[1]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[1]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[2]  Number(10,2)    �ֶ�2
    'Ld_bckbje[2]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[2]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[3]  Number(10,2)    �ֶ�3
    'Ld_bckbje[3]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[3]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[4]  Number(10,2)    �ֶ�4
    'Ld_bckbje[4]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[4]    Number(10,2)    ����ʵ�����
    'Ld_bcfd[5]  Number(10,2)    �ֶ�5
    'Ld_bckbje[5]    Number(10,2)    ���οɱ����
    'Ld_fdbxje[5]    Number(10,2)    ����ʵ�����
    
    dbl���� = Val(Format(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(5), "=")(1), "#0.00"))
    dblҽ������ = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(6), "=")(1))
    dbl�ʻ�֧�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(7), "=")(1))
    dbl�ֽ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(8), "=")(1))
    str������ˮ�� = Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(12), "=")(1)
    
    dbl���οɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(13), "=")(1))
    dbl�����ۼ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(14), "=")(1))
    str�������� = Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(15), "=")(1)
    dbl�ֶ�1 = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(16), "=")(1))
    dbl�ֶ�1�ɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(17), "=")(1))
    dbl�ֶ�1ʵ�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(18), "=")(1))
    dbl�ֶ�2 = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(19), "=")(1))
    dbl�ֶ�2�ɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(20), "=")(1))
    dbl�ֶ�2ʵ�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(21), "=")(1))
    dbl�ֶ�3 = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(22), "=")(1))
    dbl�ֶ�3�ɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(23), "=")(1))
    dbl�ֶ�3ʵ�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(24), "=")(1))
    dbl�ֶ�4 = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(25), "=")(1))
    dbl�ֶ�4�ɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(26), "=")(1))
    dbl�ֶ�4ʵ�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(27), "=")(1))
    dbl�ֶ�5 = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(28), "=")(1))
    dbl�ֶ�5�ɱ� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(29), "=")(1))
    dbl�ֶ�5ʵ�� = Val(Split(Split(gstrOutput_��Ϫũҽ, mstrSplit)(30), "=")(1))
    
    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ϫũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_��Ϫũҽ.�ܷ��� & "," & dbl�ֽ� & ",0," & dbl���� & "," & dblҽ������ & ",0,0," & _
        dbl�ʻ�֧�� & ",'" & str������ˮ�� & "'," & lng��ҳID & ",null,'" & GetKey(lng����ID, lng��ҳID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    gstrSQL = "ZL_���㸽����Ϣ_INSERT(" & lng����ID & ",'" & str������ˮ�� & "'," & dbl���οɱ� & "," & dbl�����ۼ� & "," & _
        "'" & str�������� & "'," & dbl�ֶ�1 & "," & dbl�ֶ�1�ɱ� & "," & dbl�ֶ�1ʵ�� & "," & _
        dbl�ֶ�2 & "," & dbl�ֶ�2�ɱ� & "," & dbl�ֶ�2ʵ�� & "," & _
        dbl�ֶ�3 & "," & dbl�ֶ�3�ɱ� & "," & dbl�ֶ�3ʵ�� & "," & _
        dbl�ֶ�4 & "," & dbl�ֶ�4�ɱ� & "," & dbl�ֶ�4ʵ�� & "," & _
        dbl�ֶ�5 & "," & dbl�ֶ�5�ɱ� & "," & dbl�ֶ�5ʵ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㸽����Ϣ")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʼ�¼�����ϴ���־")

    סԺ����_��Ϫũҽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��Ϫũҽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    Dim StrInput As String
    Dim strCardNO As String
    Dim lng����ID As Long
    Dim lng����ID As Long, lng��ҳID As Long, lng��ҳID_��ǰ As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsBalance As New ADODB.Recordset
    On Error GoTo errHand

    'ȡ����ID
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!ID

    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=[1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ˮ��", lng����ID)
    If rsBalance.RecordCount = 0 Then
        MsgBox "û���ҵ�ԭʼ�����¼���޷�����סԺ���������", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsBalance!����ID
    lng��ҳID = rsBalance!��ҳID
    
    'ȡ��ǰ��ҳID
    gstrSQL = "Select Nvl(סԺ����,0) AS ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID_��ǰ = rsTemp!��ҳID
    
    If lng��ҳID <> lng��ҳID_��ǰ Then
        MsgBox "���ܳ����ϴ�סԺ�ڼ�Ľ��㵥�����ȳ���������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ��������
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", TYPE_��Ϫũҽ, lng����ID)
    strCardNO = rsTemp!����

    '���ý������
'    Jylx    varchar(1)  ��ҽ����(=0 �������� = 1 ȡ����Ժ)
'    Jyhm    varchar(20) ��ҽ����
'    Kzhm    varchar(30) ��֤����
'    Tfrq    Datetime    ��������(�˷�����) (�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)
'    Czy Varchar(10) ����Ա����
    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & rsBalance!��ע & mstrSplit & _
        "Kzhm=" & strCardNO & mstrSplit & "Tfrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & _
       "Czy=" & UserInfo.����
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_BalanceCancel, StrInput)
    If Not ���ýӿ�_��Ϫũҽ() Then Exit Function

    '���汾�ν������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ϫũҽ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsBalance!�������ý��, 0) & "," & -1 * Nvl(rsBalance!ȫ�Ը����, 0) & "," & -1 * Nvl(rsBalance!�����Ը����, 0) & "," & -1 * Nvl(rsBalance!����ͳ����, 0) & "," & -1 * Nvl(rsBalance!ͳ�ﱨ�����, 0) & ",0,0," & _
        -1 * Nvl(rsBalance!�����ʻ�֧��, 0) & ",null,null,null,'" & Nvl(rsBalance!��ע) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
    
    gstrSQL = "Select * From ���㸽����Ϣ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���㸽����Ϣ", lng����ID)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "ZL_���㸽����Ϣ_INSERT(" & lng����ID & ",'" & Nvl(rsTemp!���㵥��) & "'," & -1 * Nvl(rsTemp!���οɱ����, 0) & "," & -1 * Nvl(rsTemp!�����ۼƿɱ����, 0) & "," & _
            "'" & Nvl(rsTemp!��������) & "'," & -1 * Nvl(rsTemp!�ֶ�1, 0) & "," & -1 * Nvl(rsTemp!�ֶ�1�ɱ�, 0) & "," & -1 * Nvl(rsTemp!�ֶ�1ʵ��, 0) & "," & _
            -1 * Nvl(rsTemp!�ֶ�2, 0) & "," & -1 * Nvl(rsTemp!�ֶ�2�ɱ�, 0) & "," & -1 * Nvl(rsTemp!�ֶ�2ʵ��, 0) & "," & _
            -1 * Nvl(rsTemp!�ֶ�3, 0) & "," & -1 * Nvl(rsTemp!�ֶ�3�ɱ�, 0) & "," & -1 * Nvl(rsTemp!�ֶ�3ʵ��, 0) & "," & _
            -1 * Nvl(rsTemp!�ֶ�4, 0) & "," & -1 * Nvl(rsTemp!�ֶ�4�ɱ�, 0) & "," & -1 * Nvl(rsTemp!�ֶ�4ʵ��, 0) & "," & _
            -1 * Nvl(rsTemp!�ֶ�5, 0) & "," & -1 * Nvl(rsTemp!�ֶ�5�ɱ�, 0) & "," & -1 * Nvl(rsTemp!�ֶ�5ʵ��, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������㸽����Ϣ")
    End If

    סԺ�������_��Ϫũҽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function IsYBPatient(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�ж�ָ�����˱����Ƿ���ҽ����ݾ���
    gstrSQL = " Select 1 From ������ҳ Where ����=" & TYPE_��Ϫũҽ & " And (����ID,��ҳID) IN " & _
              "     (Select ����ID,סԺ���� From ������Ϣ Where ����ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�ָ�����˱����Ƿ���ҽ����ݾ���", lng����ID)
    IsYBPatient = (rsTemp.RecordCount <> 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetKey(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim strKey As String
    '�����ҳIDС��һλ����ǰ���һ����
    strKey = lng��ҳID
    If Len(strKey) = 1 Then strKey = "0" & strKey
    GetKey = lng����ID & strKey
End Function

Private Sub UploadDetail(ByVal int��¼���� As Integer, ByVal int��¼״̬ As Integer, ByVal strNO As String, ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select NO,���,��¼����,��¼״̬ From סԺ���ü�¼" & _
              " Where ��¼����=[1] And ��¼״̬=[2] And NO=[3] And ����ID=[4] And Nvl(�Ƿ��ϴ�,0)=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵Ĵ�����ϸ", int��¼����, int��¼״̬, strNO, lng����ID)
    
    '��ָ�����˵Ĵ�����ϸ���ϴ����
    With rsTemp
        Do While Not .EOF
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���ϴ���־")
            .MoveNext
        Loop
    End With
End Sub

Public Function ������Ժ��Ϣ_��Ϫũҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln��Ժ As Boolean = False) As Boolean
    Dim StrInput As String
    Dim strKey As String                    '��Ժ�ǼǺ�
    Dim strCardNO As String, lngDisease As Long '���˵�ҽ�ƿ��ż�����ID
    Dim strRegistCode As String             '�Һŵ���
    Dim strInHospitalDate As String         '��Ժ����
    Dim strRegisterOffice As String         '�������
    Dim strDiseaseCode As String            '���ִ���
    Dim strDiagnose As String               '��Ժ���
    Dim strRegisterDoctor As String         'ҽ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strKey = GetKey(lng����ID, lng��ҳID)
    'ȡ���˵�ҽ�ƿ���
    gstrSQL = "Select ����,Nvl(����ID,0) ����ID From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�ҽ�ƿ���", TYPE_��Ϫũҽ, lng����ID)
    strCardNO = rsTemp!����
    lngDisease = rsTemp!����ID
    
    'ȡ������ҽ��
    gstrSQL = " Select A.��Ժ����,B.���� ����,A.סԺҽʦ ҽ�� From ������ҳ A,���ű� B " & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A." & IIf(bln��Ժ = False, "��Ժ����ID", "��Ժ����ID") & "=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ҽ��", lng����ID, lng��ҳID)
    strInHospitalDate = Format(rsTemp!��Ժ����, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!ҽ��)
    strRegisterOffice = Nvl(rsTemp!����)
    'ȡ���ִ���
    gstrSQL = "Select ���� From ��������Ŀ¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ִ���", lngDisease)
    If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!����
    'ȡ��Ժ���
    strDiagnose = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, False)
    
    '��Σ�
'    Rydjh   varchar(20) ��Ժ�ǼǺţ���ҽԺ���ݿ��е�Ψһ����
'    Kzhm    varchar(30) ��֤����
'    Jzks    varchar(20) �������
'    Ysxm    varchar(10) ҽ������
'    Bzdm    Varchar(20) ���ִ���
'    Ryzdsm  varchar(254)    ��Ժ���˵��
'    Ryrq    Datetime    ��Ժ����(�̶���ʽ19λ��yyyy-mm-dd hh:mm:ss)��������ͬ
'    Czy Varchar(10) ����Ա����
    '���ز���:
'    Returncode  Long    �������0��ʾ�ɹ�
'    Returninfo  varchar(50) ��Ӧ�Ĵ�����ʾ
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & _
        "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
        "Bzdm=" & strDiseaseCode & mstrSplit & "Ryzdsm=" & strDiagnose & mstrSplit & _
        "Ryrq=" & strInHospitalDate & mstrSplit & "Czy=" & UserInfo.����
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_ModifyInfo, StrInput)
    If Not ���ýӿ�_��Ϫũҽ() Then Exit Function
    
    ������Ժ��Ϣ_��Ϫũҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ���²���_��Ϫũҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim lng����ID As Long
    lng����ID = frm����ѡ��_��Ϫũҽ.ChooseDisease(lng����ID)
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ϫũҽ & ",'����ID','''" & lng����ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���没��ID")
    
    Call ������Ժ��Ϣ_��Ϫũҽ(lng����ID, lng��ҳID)
End Sub

Public Function getС����ҩ����(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long, ByVal str���� As String) As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡС����ҩ����
    getС����ҩ���� = str����
    
    '�жϱ�����Ժ�Ƿ���С����ҩ�ķ�ʽ��Ժ
    gstrSQL = "Select Nvl(С����ҩ,0) AS С����ҩ From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�жϱ�����Ժ�Ƿ���С����ҩ�ķ�ʽ��Ժ", lng����ID, TYPE_��Ϫũҽ)
    If rsTemp.RecordCount = 0 Then Exit Function
    If rsTemp!С����ҩ = 0 Then Exit Function
    
    '��ȡС����ҩ���벢����
    gstrSQL = "Select ��ע From ����֧����Ŀ Where ����=[1] ANd �շ�ϸĿID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡС����ҩ����", TYPE_��Ϫũҽ, lng�շ�ϸĿID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If InStr(1, Nvl(rsTemp!��ע), "|") = 0 Then Exit Function
    getС����ҩ���� = Split(rsTemp!��ע, "|")(0)
    If getС����ҩ���� = "" Then getС����ҩ���� = str����
End Function
