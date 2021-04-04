Attribute VB_Name = "mdlDYEYMZYF"
Option Explicit

Public gobjSOAP As Object  '�ӿڶ���
Public gstrIP As String    '����ip
Public gblnShowMsg As Boolean   '�Ƿ񵯳��Ի�����ʾ�������շ���Ҫ��
Public gstrUnit As String   '�û�ע���û���
Public gstrOutPut As String     '��־�������
Public gblnUpdateFlag As Boolean

Public Const GCST_UNIT_DYEY = "����ҽ�ƴ�ѧ�����ڶ�ҽԺ"
Public Const GCST_UNIT_YZSZYY = "��������ҽԺ"
Public Const GCST_UNIT_JLSZXYY = "����������ҽԺ"
Public Const GCST_UNIT_CQFLQZYY = "�����и�������ҽԺ"
Public Const GCST_UNIT_YNYXRMYY = "����ʡ��Ϫ������ҽԺ"
Public Const GCST_UNIT_YQMY = "��Ȫúҵ�����ţ��������ι�˾��ҽԺ"
Public Const GCST_UNIT_BTSZXYY = "��ͷ������ҽԺ"

Public Const GINT_SEND_TYPE = 1           '0-����ʼ��ҩ���̣�1-�п�ʼ��ҩ��������ҩ����
Public Const GINT_STARTSEND_TYPE = 1      '0-��ť��ʽ��ʼ��ҩ��1-ˢ����ʽ��ʼ��ҩ
Public Const GBLN_OUTPUTLOG_DETAIL = True   'д��־ʱ�Ƿ������ϸ����(�ϴ����Է��ӿڵ���ϸ����)�������falseֻ�ڳ���ʱ�����ϸ����

'�̶�ҩ��
Public Const GCST_DRUGID_DYEY = 176         '����ҽ�ƴ�ѧ�����ڶ�ҽԺ������ҩ��

Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(5) As IPINFO  'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum gType
    IntDrug = 101       '�ϴ�ҩƷ��������
    IntStore = 102      '�ϴ�ҩƷ�������
    IntDept = 104       '�ϴ���������
    IntDetail = 201     '�ϴ�������ϸ
    IntStartList = 202  '�ϴ�������������ʼ��ҩ
    IntEndList = 203    '�ϴ�����������������ҩ
    IntReturnAll = 205  '�����˷ѣ�ȫ��ģʽ
End Enum

Private mstrSQL As String

Private mobjFSO As New FileSystemObject

Public Function DYEY_MZ_TransData(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, _
    ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, Optional ByVal strNO As String, _
    Optional ByVal lngStockID As Long) As Boolean
'1.��WebService��������
'2.���ӿں�������
    Dim i As Integer
    Dim intRetval As Integer
    Dim strRETMSG As String
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    Dim strOutput As String
    
    On Error GoTo errHandle
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "�����ϴ���", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�����ϴ���"
        End If
    End If
    If gstrIP = "" Then
        gstrIP = GetLocalIP
    End If
    
    For i = 0 To UBound(arrXML)
        If gobjSOAP.TransConsisData(intOprId, intType, CStr(arrXML(i)), gstrIP, strUserCode, strUserName, intRetval, strRETMSG) <> 1 Then
            If gblnShowMsg Then
                MsgBox strRETMSG, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strRETMSG
            End If
            
            '��־���
            strOutput = "SOAP.TransConsisData��" & vbNewLine & _
                        "   XML��" & arrXML(i) & vbNewLine & _
                        "   intRetval��" & intRetval & vbNewLine & _
                        "   strRETMSG��" & strRETMSG & vbNewLine
            Call OutputLog(strOutput)
            
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
        ElseIf intType = gType.IntDetail Then
            '���ϴ���Ϣ��ȡҩ��ID
            lngDrugStockID = GetStockID(arrXML(i))
            
            If Not SetSendWin(lngDrugStockID, strNO, intRetval) Then
                If gblnShowMsg Then
                    MsgBox "���������ķ�ҩ����ʧ�ܣ�", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "���������ķ�ҩ����ʧ�ܣ�"
                End If
            End If
        End If
    Next
    
    DYEY_MZ_TransData = True
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "�ϴ���ɣ�", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�ϴ���ɣ�"
        End If
    End If
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
End Function

Public Function DYEY_MZ_TransData_CQFLQZYY(ByVal intType As Integer, ByVal intOprId As Integer, ByVal strUserCode As String, _
    ByVal strUserName As String, ByVal arrXML As Variant, ByRef strReturn As String, ByRef strOutput As String, _
    Optional ByVal strNO As String, Optional ByVal lngStockID As Long) As Boolean
    
'1.��WebService��������
'2.���ӿں�������
'3.���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    
    Dim i As Integer
    Dim blnShow As Boolean
    Dim lngDrugStockID As Long
    Dim strXML_In As String
    Dim strXML_Out As String
    Dim strOut_RETVAL As String
    Dim strOut_RETMSG As String
    Dim strOut_RETCODE As String
    Dim strTmp As String
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����DYEY_MZ_TransData_CQFLQZYY"
    strOutput = strOutput & vbCrLf & "ҵ����룺" & intType
    
    If intType = gType.IntDrug Or intType = gType.IntStore Then
        If gblnShowMsg Then
            MsgBox "�����ϴ���", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�����ϴ���"
        End If
    End If
    
    If gstrIP = "" Then
        gstrIP = GetLocalIP
        
        strOutPutExeStep = "ȡ�ͻ���IP"
    End If

    For i = 0 To UBound(arrXML)
        'XML��ʼ
        strXML_In = "<ROOT>"
        
        'ҵ�������Ϣ
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPSYSTEM", "HIS")
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPWINID", IIf(intOprId = 0, "", intOprId))
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPTYPE", intType)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPIP", gstrIP)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPMANNO", strUserCode)
        strXML_In = strXML_In & vbCrLf & GetXMLFormat("OPMANNAME", strUserName)
    
        'ҵ���������Ϣ
        strXML_In = strXML_In & vbCrLf & CStr(arrXML(i))
        
        'XML������־
        strXML_In = strXML_In & vbCrLf & "</ROOT>"
        
        strOutPutExeStep = "�ϴ�����" & i + 1 & "/" & UBound(arrXML) + 1
        
        '����ϴ�����
        If GBLN_OUTPUTLOG_DETAIL = True Then strOutput = strOutput & vbCrLf & strXML_In
       
        '���ýӿڷ����ϴ�����
        strXML_Out = gobjSOAP.HisTransData(strXML_In)
        
        strOutPutExeStep = "���öԷ��ӿ����"
        
        strOutput = strOutput & vbCrLf & "������Ϣ" & vbCrLf & strXML_Out
        
        'ȥ���س����з�
        strXML_Out = Replace(strXML_Out, vbCrLf, "")
        strXML_Out = Replace(strXML_Out, vbCr, "")
        strXML_Out = Replace(strXML_Out, vbLf, "")
        
        '�������ز���
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETVAL>") - 1)
        strOut_RETVAL = Mid(strTmp, InStr(1, strTmp, "<RETVAL>") + Len("<RETVAL>"))
        
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETMSG>") - 1)
        strOut_RETMSG = Mid(strTmp, InStr(1, strTmp, "<RETMSG>") + Len("<RETMSG>"))
        
        strTmp = strXML_Out
        strTmp = Mid(strTmp, 1, InStr(1, strTmp, "</RETCODE>") - 1)
        strOut_RETCODE = Mid(strTmp, InStr(1, strTmp, "<RETCODE>") + Len("<RETCODE>"))
               
        strOutPutExeStep = "�������ز������"
               
        '����1��ʾ�ӿڵ��óɹ�������ֵΪ���ɹ�
        If strOut_RETCODE <> "1" Then
            If gblnShowMsg Then
                MsgBox strOut_RETMSG, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strOut_RETMSG
            End If
            If blnShow Then frmDYEY_MZ_TransDrug.UnloadMe
            
            strOutput = strOutput & vbCrLf & "�ϴ����ݴ���"
            If GBLN_OUTPUTLOG_DETAIL = False Then strOutput = strOutput & vbCrLf & "���һ���ϴ�����" & vbCrLf & CStr(arrXML(i))
            strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�DYEY_MZ_TransData_CQFLQZYY"
            Call OutputLog(strOutput)
    
            Exit Function
        End If
        
        If intType = gType.IntDrug Or intType = gType.IntStore Or intType = gType.IntDept Then
            If i = 0 Then
                frmDYEY_MZ_TransDrug.Show
                blnShow = True
                
                strOutPutExeStep = "���ϴ�ҩƷ��Ϣ����"
            End If
            
            Call frmDYEY_MZ_TransDrug.ChangePrg(i + 1, UBound(arrXML) + 1, intType)
            
            strOutPutExeStep = "�ϴ�ҩƷ��Ϣ���ڽ�����ִ��"
        ElseIf intType = gType.IntDetail Then
            '�����ϴ���Ϣȡ�ⷿID
            lngDrugStockID = GetStockID(arrXML(i))
            
            strOutPutExeStep = "ȡ�ⷿID"
            
            If Not SetSendWin(lngDrugStockID, strNO, Val(strOut_RETMSG)) Then
                If gblnShowMsg Then
                    MsgBox "���������ķ�ҩ����ʧ�ܣ�", vbCritical, GSTR_MESSAGE
                Else
                    strReturn = "���������ķ�ҩ����ʧ�ܣ�"
                End If
                
                strOutPutExeStep = "���������ķ�ҩ����ʧ�ܣ���" & lngDrugStockID & "��" & strNO & "��" & strOut_RETMSG & "��"
                strOutput = strOutput & vbCrLf & "��������ʧ�ܣ����ⷿ��" & lngDrugStockID & "��NO��" & strNO & "��������Ϣ��" & strOut_RETMSG & "��"
            Else
                strOutPutExeStep = "���������ķ�ҩ���ڳɹ�����" & lngDrugStockID & "��" & strNO & "��" & strOut_RETMSG & "��"
                strOutput = strOutput & vbCrLf & "�������ڳɹ������ⷿ��" & lngDrugStockID & "��NO��" & strNO & "��������Ϣ��" & strOut_RETMSG & "��"
            End If
            
        End If
    Next
    
    If intType = gType.IntDrug Or intType = gType.IntStore Or intType = gType.IntDept Then
        If gblnShowMsg Then
            MsgBox "�ϴ���ɣ�", vbInformation, GSTR_MESSAGE
        Else
            strReturn = "�ϴ���ɣ�"
        End If
    End If
    
    DYEY_MZ_TransData_CQFLQZYY = True
        
    strOutput = strOutput & vbCrLf & "ִ�гɹ���DYEY_MZ_TransData_CQFLQZYY"
   
    Exit Function
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�DYEY_MZ_TransData_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXML_Drug() As Variant
'��ҩƷ������Ϣ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strErrMsg As String
    
    On Error GoTo errHandle
'    MsgBox "��ȡ����"
    strErrMsg = "��ȡ����"
    mstrSQL = "Select Distinct a.id ҩƷ���, a.���� ҩƷ����, e.���� ҩƷ��Ʒ��, a.��� ҩƷ���, a.��� ҩƷ��װ���, b.���ﵥλ ҩƷ��λ," & vbNewLine & _
              "    round(b.ҩ���װ/b.�����װ, 2) ��װ��,b.����ɷ����,a.���� ҩƷ����, c.�ּ� * b.�����װ ҩƷ�۸�, d.ҩƷ����, " & vbNewLine & _
              "    b.�����װ, a.����ʱ�� ������ʱ��, f.���� ҩƷƴ��, d.������� " & vbNewLine & _
              "From �շ���ĿĿ¼ a, ҩƷ��� b, �շѼ�Ŀ c, ҩƷ���� d, �շ���Ŀ���� e, �շ���Ŀ���� f " & vbNewLine & _
              "Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid And b.ҩ��id = d.ҩ��id And a.Id = e.�շ�ϸĿid(+) And a.Id = f.�շ�ϸĿid(+) And " & vbNewLine & _
              "    (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And Sysdate Between c.ִ������ And " & vbNewLine & _
              "    Nvl(c.��ֹ����, Sysdate) And e.����(+) = 3 And f.����(+) = 1 And f.����(+) = 1"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Drug")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Drug")
    End If
    strErrMsg = "���ݻ�ȡ���"
    strXML = ""
    arrXML = Array()
    
    strErrMsg = "XML��ʼ"
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_NAME = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "TRADE_NAME = """ & SpecialChar(!ҩƷ��Ʒ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SPEC = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PACKAGE = """ & NVL(!�����װ) & """"  ' & SpecialChar(!ҩƷ��װ���) & """"
                strDrug = strDrug & vbCrLf & "DRUG_UNIT = """ & SpecialChar(!ҩƷ��λ) & """"
                strDrug = strDrug & vbCrLf & "FIRM_ID = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "DRUG_PRICE = """ & NVL(!ҩƷ�۸�) & """"
                strDrug = strDrug & vbCrLf & "DRUG_FORM = """ & SpecialChar(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "DRUG_SORT = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "BARCODE = """""
                strDrug = strDrug & vbCrLf & "LAST_DATE = """ & Format(!������ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PINYIN = """ & SpecialChar(!ҩƷƴ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_CONVERTATION = """ & NVL(!��װ��) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                If Len(strXML & strDrug) > 3900 Then
                    '����ǰ����ӵ�����
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "װ������1"
                    '����ƴ���µ�XML
                    strXML = strTitle & vbCrLf & strDrug
                Else
                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                End If
                
                rsTemp.MoveNext
                If .EOF And strXML <> "" Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    strErrMsg = "װ������2"
                End If
            Loop
        End If
    End With
    
    strErrMsg = "��ȡ����"
    GetXML_Drug = arrXML
    strErrMsg = "��������"
    Exit Function

errHandle:
    Debug.Print strErrMsg
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Drug_CQFLQZYY(ByRef strOutput As String) As Variant
'��ҩƷ������Ϣ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim lngCount As Long
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    '�ӿ����ݸ�ʽ
    '�ֶ���      ����            ˵��       NULL
    'Drug_code   Nvarchar(200)   ҩƷ���    N
    'Drug_name   Nvarchar(200)   ҩƷ����    N
    'Tradename   Nvarchar(200)   ��Ʒ����    Y
    'Englishname Nvarchar(200)   ҩƷӢ����  Y
    'Pinyin  Nvarchar(1000)  ��Ʒƴ����  Y
    'SortType1   Nvarchar(40)    ҩƷ���    Y
    'SortType2   Nvarchar(40)    ҩƷ����    Y
    'Drug_spec   Nvarchar(200)   ҩƷ���    N
    'MinSpecs    Nvarchar(200)   ҩƷ��С���    Y
    'Unit    Nvarchar(40)    ��װ��λ    N
    'MaxUNIT Nvarchar(40)    ���װ��λ  N
    'MinUNIT Nvarchar(40)    ��С��λ    N
    'Dosage  Numeric(20,6)   ��С��λ����    N
    'DosageUnit  Nvarchar(40)    ������λ    Y
    'Price1  Numeric(20,6)   ҩƷ�۸�    N
    'Convertion1 Numeric(10,0)   ���װ��λ����װ��λ������  N
    'Convertion2 Numeric(10,0)   ��װ��λ��С��װ��λ������  N
    'Firm_id Nvarchar(200)   �������ұ���    Y
    'Firm_name   Nvarchar(200)   ������������    Y
    'Passno  Nvarchar(200)   ��׼�ĺ�/ע��֤��   Y
    'BarCode Nvarchar(200)   ҩƷ����    Y
    'StorageCondition    Nvarchar(200)   ��������    Y
    'Storagetype Char(1) ��������(Ĭ��'0')   N
    'Allowind    Char(1) ͣ�ñ�־��Y/N��
    '                       'Y'����
    '                       'N'ͣ�� N
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_Drug_CQFLQZYY"
    
    On Error GoTo errHandle
              
    mstrSQL = "Select Distinct a.���� As ҩƷ���, a.���� As ҩƷ����, e.���� As ҩƷ��Ʒ��, g.���� As Ӣ����, f.���� As ƴ����," & vbNewLine & _
        " Decode(a.���, 5, '��ҩ', 6, '��ҩ', '��ҩ') As ���, d.ҩƷ����, a.��� As ҩƷ���, b.���ﵥλ As ��װ��λ, b.ҩ�ⵥλ As ���װ��λ," & vbNewLine & _
        " a.���㵥λ As ��С��λ, b.����ϵ��, i.���㵥λ As ������λ, c.�ּ� * b.�����װ As ҩƷ�۸�, Round(b.ҩ���װ / b.�����װ, 2) As ���װϵ��," & vbNewLine & _
        " b.�����װ As ��װϵ��, j.���� As �������ұ���, j.���� As ������������, b.�ϴ���׼�ĺ� As ��׼�ĺ�," & vbNewLine & _
        " Decode(Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')), To_Date('3000-01-01', 'yyyy-MM-dd'), 'Y', 'N') As ͣ�ñ�־" & vbNewLine & _
        " From �շ���ĿĿ¼ A, ҩƷ��� B, �շѼ�Ŀ C, ҩƷ���� D, �շ���Ŀ���� E, �շ���Ŀ���� F, �շ���Ŀ���� G, ������ĿĿ¼ I, ҩƷ������ J" & vbNewLine & _
        " Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid And b.ҩ��id = d.ҩ��id And a.Id = e.�շ�ϸĿid(+) And a.Id = f.�շ�ϸĿid(+) And" & vbNewLine & _
        " Sysdate Between c.ִ������ And Nvl(c.��ֹ����, Sysdate) And e.����(+) = 3 And f.����(+) = 1 And f.����(+) = 1 And" & vbNewLine & _
        " a.Id = g.�շ�ϸĿid(+) And g.����(+) = 2 And b.ҩ��id = i.Id And b.�ϴβ��� = j.����(+)"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Drug_CQFLZYY")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Drug_CQFLZYY")
    End If
    
    strOutPutExeStep = "ִ��SQL�ɹ�"
    
    strXML = ""
    arrXML = Array()

    With rsTemp
        strOutPutExeStep = "ƴװXML��begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_DRUGSVW>"
               
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(!ҩƷ���))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(!ҩƷ����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("TRADENAME", SpecialChar(!ҩƷ��Ʒ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ENGLISHNAME", SpecialChar(!Ӣ����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PINYIN", SpecialChar(!ƴ����))
                
                strOutPutExeStep = "ƴװXML��1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("SORTTYPE1", SpecialChar(!���))
                strDrug = strDrug & vbCrLf & GetXMLFormat("SORTTYPE2", SpecialChar(!ҩƷ����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(!ҩƷ���))
                strDrug = strDrug & vbCrLf & GetXMLFormat("MINSPECS", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("UNIT", SpecialChar(NVL(!��װ��λ)))
                
                strOutPutExeStep = "ƴװXML��2"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("MAXUNIT", SpecialChar(NVL(!���װ��λ)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("MINUNIT", SpecialChar(NVL(!��С��λ)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DOSAGE", NVL(!����ϵ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DOSAGEUNIT", SpecialChar(NVL(!������λ)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRICE1", NVL(!ҩƷ�۸�))
                
                strOutPutExeStep = "ƴװXML��3"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("CONVERTION1", NVL(!���װϵ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("CONVERTION2", NVL(!��װϵ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(NVL(!�������ұ���)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(NVL(!������������)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PASSNO", SpecialChar(NVL(!��׼�ĺ�)))
                
                strOutPutExeStep = "ƴװXML��4"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("BARCODE", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("STORAGECONDITION", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("STORAGETYPE", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("ALLOWIND", NVL(!ͣ�ñ�־))
                
                strOutPutExeStep = "ƴװXML��5"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_DRUGSVW>"
                
                lngCount = lngCount + 1
                
                'ÿ500��ҩƷ���һ����������ϴ�
                If lngCount > 500 Then
                    '����ǰ����ӵ�����
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    '����ƴ���µ�XML
                    strXML = strDrug
                    lngCount = 0
                    
                    strOutPutExeStep = "ƴװXML��7"
                Else
                    strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                    
                    strOutPutExeStep = "ƴװXML��6"
                End If
                
                rsTemp.MoveNext
                
                If .EOF And strXML <> "" Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    
                    strOutPutExeStep = "ƴװXML��end"
                End If
            Loop
        End If
    End With
    
    GetXML_Drug_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_Drug_CQFLQZYY"
  
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_Drug_CQFLQZYY"
    Call OutputLog(strOutput)
End Function

Public Function GetXML_RecipeDetail(ByVal strStockIDs As String, ByVal strNO As String) As Variant
'��������ϸ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    
    Call OutputLog("����GetXML_RecipeDetail")
    
    On Error GoTo errHandle
    '��ȡ��������Ϣ
    strSQL = "Select a.�������� ����ʱ��, a.����, a.No �������, a.�ⷿid ��ҩҩ��, c.����id ���￨��, a.���� ��������, Decode(a.���ȼ�, 1, '01', '00') ��������, " & vbNewLine & _
             "    c.�������� ���߳�������, c.�Ա� �����Ա�, c.��� �������, c.ҽ�Ƹ��ʽ ҽ������, Sum(d.Ӧ�ս��) ����, Sum(d.ʵ�ս��) ʵ������," & vbNewLine & _
             "    f.id ��������, d.������ ����ҽ��, d.������ ¼����, Decode(a.���ȼ�, 1, '1', '2') ��ҩ���ȼ� " & vbNewLine & _
             "From δ��ҩƷ��¼ a, ������Ϣ c, ������ü�¼ d, ҩƷ�շ���¼ e, ���ű� f " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I " & vbNewLine & _
             "Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id And e.����id = d.Id And " & vbNewLine & _
             "    d.��������id = f.Id And a.���� = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.�ⷿid || ';') > 0 ")

    strSQL = strSQL & _
             "Group By a.��������, a.����, a.No, a.�ⷿid, c.����id, a.����, Decode(a.���ȼ�, 1, '01', '00'), c.��������, c.�Ա�, " & vbNewLine & _
             "    c.���, c.ҽ�Ƹ��ʽ, f.id, d.������,d.������, Decode(a.���ȼ�, 1, '1', '2') "

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
    mstrSQL = "select * from (" & mstrSQL & ") Order By ��ҩҩ��, ���￨�� "
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    End If
    
    '��ȡ������ϸ��Ϣ
    strSQL = "Select Distinct a.��������, a.����, a.No, a.���, b.id ҩƷ����, b.���� ҩƷ����, c.���� ҩƷ��Ʒ��, b.��� ҩƷ���, b.��� ҩƷ��װ���, " & vbNewLine & _
             "    d.���ﵥλ ҩƷ��λ, a.���� ҩƷ����, a.���ۼ� * d.�����װ ҩƷ�۸�, a.ʵ������ / d.�����װ ����, e.Ӧ�ս�� ����,e.����id," & vbNewLine & _
             "    e.ʵ�ս�� ʵ������, a.���� ҩƷ����, a.�ⷿid, a.�÷�, f.ִ��Ƶ��, g.���㵥λ ������λ " & vbNewLine & _
             "From ҩƷ�շ���¼ a, �շ���ĿĿ¼ b, �շ���Ŀ���� c, ҩƷ��� d, ������ü�¼ e, ����ҽ����¼ f, ������ĿĿ¼ g " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I " & vbNewLine & _
             "Where a.ҩƷid = b.Id And a.ҩƷid = c.�շ�ϸĿid(+) And a.ҩƷid = d.ҩƷid And a.����id = e.Id and d.ҩ��id=g.id " & vbNewLine & _
             "    And e.ҽ����� = f.Id(+) And c.����(+) = 3 And a.���� = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.�ⷿid || ';') > 0 ")

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
Call OutputLog("��ѯ������ϸ��ʼ")
    If gintMode = 0 Then
        Set rsDetails = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    Else
        Set rsDetails = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail", ";" & strStockIDs & ";", strNO)
    End If
    strXML = ""
    arrXML = Array()
    
Call OutputLog("��ѯ������ϸ��ɡ�")
    
'    '�ⷿIDΪ0�����������������
'    If lngStockID = 0 Then
'Call OutputLog("ִ��GetXML_RecipeDetailEx")
'        If GetXML_RecipeDetailEx(rsTemp, rsDetails, arrXML) Then
'            GetXML_RecipeDetail = arrXML
'        End If
'        Exit Function
'    End If
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_ID = """ & NVL(!���￨��) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!��������) & """"
                strDrug = strDrug & vbCrLf & "PATIENT_TYPE = """ & NVL(!��������) & """"
                strDrug = strDrug & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "SEX = """ & SpecialChar(!�����Ա�) & """"
                strDrug = strDrug & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!�������) & """"
                strDrug = strDrug & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!ҽ������) & """"
                strDrug = strDrug & vbCrLf & "PRESC_ATTR = """""
                strDrug = strDrug & vbCrLf & "PRESC_INFO = """""
                strDrug = strDrug & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!�������))
                strDrug = strDrug & vbCrLf & "RCPT_REMARK = """""
                strDrug = strDrug & vbCrLf & "REPETITION = ""1"""
                strDrug = strDrug & vbCrLf & "COSTS = """ & NVL(!����) & """"
                strDrug = strDrug & vbCrLf & "PAYMENTS = """ & NVL(!ʵ������) & """"
                strDrug = strDrug & vbCrLf & "ORDERED_BY = """ & NVL(!��������) & """"
                strDrug = strDrug & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!����ҽ��) & """"
                strDrug = strDrug & vbCrLf & "ENTERED_BY = """ & SpecialChar(!¼����) & """"
                strDrug = strDrug & vbCrLf & "DISPENSE_PRI = """ & NVL(!��ҩ���ȼ�) & """"
                strDrug = strDrug & vbCrLf & ">"
                
                '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
                rsDetails.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��)
                rsDetails.Sort = "���"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW"
                    strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetails!��������, "yyyy-MM-DDThh:mm:ss") & """"
                    strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetails!no) & """"
                    strDetail = strDetail & vbCrLf & "ITEM_NO = """ & NVL(rsDetails!���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetails!ҩƷ��Ʒ��) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetails!ҩƷ���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetails!ҩƷ��װ���) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetails!ҩƷ��λ) & """"
                    strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetails!ҩƷ�۸�) & """"
                    strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetails!����) & """"
                    strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetails!����) & """"
                    strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetails!ʵ������) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetails!ҩƷ����) & """"
                    strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetails!������λ) & """"
                    strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetails!�÷�) & """"
                    strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetails!ִ��Ƶ��) & """"
                    strDetail = strDetail & vbCrLf & ">"
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
Call OutputLog(strXML)
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail = arrXML
    Exit Function
    
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        'MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        Call OutputLog("�������ֵ��" & strDrug & vbCr & strDetail)
    End If
End Function

Public Function GetXML_RecipeDetail_CQFLQZYY(ByVal strStockIDs As String, ByVal strNO As String, ByRef strOutput As String) As Variant
'��������ϸ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim strSQL As String
    Dim i As Integer
    Dim rsDetails As Recordset
    Dim strDetail As String
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    '��������
'    Ӣ�ı�ʶ    ���ı�ʶ    ��������    Nullable
'    Presc_date  ����ʱ��    Datetime    N
'    Presc_no    �������    Nvarchar(200)   N
'    Dispensary  ��ҩҩ�ֱ��    Nvarchar(40)    N
'    Patient_id  ���￨��    Nvarchar(40)    N
'    Patient_name    ��������    Nvarchar(200)   N
'    Invoice_no  ��Ʊ���    Nvarchar(200)   Y
'    Patient_type ��������
'    '00' ��ͨ
'    '01' ����   Nvarchar(40)    Y
'    Date_of_birth   ���߳�������    Datetime    N
'    Sex �����Ա�(��/Ů) Nvarchar(40)    N
'    Presc_identity  �������    Nvarchar(40)    Y
'    Charge_type ҽ������    Nvarchar(40)    Y
'    Presc_attr ��������
'    �ֹ���������ʱ�������ı���Ϣ    Nvarchar(1000)  Y
'    Presc_info ��������
'    ������ش��������ı���Ϣ���Ʒѷ�ʽ��    Nvarchar(1000)  Y
'    Rcpt_info   �����Ϣ    Nvarchar(1000)  Y
'    Rcpt_remark ������ע��Ϣ    Nvarchar(1000)  Y
'    Repetition  ����    Numeric(10,0)   N
'    Costs   ����    Numeric(20,6)   N
'    Payments    ʵ������    Numeric(20,6)   N
'    Ordered_by  �������ұ��    Nvarchar(40)    Y
'    Ordered_by_name ������������    Nvarchar(40)    Y
'    Prescribed_by   ����ҽ��    Nvarchar(40)    Y
'    Entered_by  ¼����  Nvarchar(40)    Y
'    Dispense_pri    ҩ���ȼ������Ѵ���ҩ�����룩���ִ�С�����ʾ    Numeric(10,0)   Y
    
    '������ϸ
'    Ӣ�ı�ʶ    ���ı�ʶ    ��������    Nullable
'    Presc_no    �������    Nvarchar(200)   N
'    Item_no ҩƷ���    Numeric(10,0)   N
'    Advice_code ҽ�����    Nvarchar(200)   Y
'    Drug_code   ҩƷ���    Nvarchar(200)   N
'    Drug_spec   ҩƷ���    Nvarchar(200)   Y
'    Drug_name   ҩƷ����    Nvarchar(200)   N
'    Firm_id ���̱��    Nvarchar(200)   Y
'    Firm_name   ��������    Nvarchar(200)   Y
'    Package_spec    ҩƷ��װ���    Nvarchar(200)   Y
'    Package_units   ҩƷ��װ��λ    Nvarchar(40)    Y
'    Quantity    ����    Numeric(20,6)   N
'    Unit    ҩƷ��λ    Nvarchar(40)    N
'    Costs   ����    Numeric(20,6)   N
'    Payments    ʵ������    Numeric(20,6)   N
'    Dosage  ҩƷ������ÿ�η�������  Nvarchar(40)    Y
'    Dosage_units    ������λ��ÿ�η��õ�λ��    Nvarchar(40)    Y
'    Administration  ҩƷ�÷���ʹ�÷�����    Nvarchar(200)   Y
'    frequency   ҩƷ������ʹ��Ƶ�� ÿ�켸�Σ�   Nvarchar(200)   Y
'    Additionusage   �����÷�    Nvarchar(200)   Y
'    Rcpt_remark ������ϸ��ע��Ϣ    Nvarchar(1000)  Y
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_RecipeDetail_CQFLQZYY"
    
    '��ȡ��������Ϣ
    strSQL = "Select a.�������� ����ʱ��, a.����, a.No �������, a.�ⷿid ��ҩҩ��, c.����id ���￨��, a.���� ��������, Decode(a.���ȼ�, 1, '01', '00') ��������, " & vbNewLine & _
             "    c.�������� ���߳�������, c.�Ա� �����Ա�, c.��� �������, c.ҽ�Ƹ��ʽ ҽ������, Sum(d.Ӧ�ս��) ����, Sum(d.ʵ�ս��) ʵ������," & vbNewLine & _
             "    f.���� �������ұ���,f.���� As ������������, d.������ ����ҽ��, d.������ ¼����, Decode(a.���ȼ�, 1, '1', '2') ��ҩ���ȼ� " & vbNewLine & _
             "From δ��ҩƷ��¼ a, ������Ϣ c, ������ü�¼ d, ҩƷ�շ���¼ e, ���ű� f " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I "
    
    '��ͷ������ҽԺҪ��סԺ������;���ҩƷ����
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " , ҩƷ��� G, ҩƷ���� T "
    End If
             
    strSQL = strSQL & "Where a.���� = e.���� And a.No = e.No And a.�ⷿid = e.�ⷿid And a.����id = c.����id And e.����id = d.Id And " & vbNewLine & _
             "    d.��������id = f.Id And a.���� = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.�ⷿid || ';') > 0 ")
    
    '��ͷ������ҽԺҪ��סԺ������;���ҩƷ����
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " And e.ҩƷid = g.ҩƷid And g.ҩ��id = t.ҩ��id And (a.���� = 9 And t.������� In ('����ҩ', '����I��') Or a.���� = 8) "
    End If
    
    strSQL = strSQL & _
             "Group By a.��������, a.����, a.No, a.�ⷿid, c.����id, a.����, Decode(a.���ȼ�, 1, '01', '00'), c.��������, c.�Ա�, " & vbNewLine & _
             "    c.���, c.ҽ�Ƹ��ʽ, f.����, f.����, d.������,d.������, Decode(a.���ȼ�, 1, '1', '2') "
             

    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
    mstrSQL = "select * from (" & mstrSQL & ") Order By ��ҩҩ��, ���￨�� "
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    End If
    
    strOutPutExeStep = "ִ�д�������ϢSQL�ɹ�"
    
    '��ȡ������ϸ��Ϣ
    strSQL = "Select Distinct a.��������, a.����, a.No, a.���, b.���� ҩƷ����, b.���� ҩƷ����, c.���� ҩƷ��Ʒ��, b.��� ҩƷ���, b.��� ҩƷ��װ���, " & vbNewLine & _
             "    d.���ﵥλ ҩƷ��λ, h.���� as �������ұ���,a.���� ������������, a.���ۼ� * d.�����װ ҩƷ�۸�, a.ʵ������ / d.�����װ ����, e.Ӧ�ս�� ����,e.����id," & vbNewLine & _
             "    e.ʵ�ս�� ʵ������,e.ҽ����� , a.���� ҩƷ����, a.�ⷿid, a.�÷�, f.ִ��Ƶ��, g.���㵥λ ������λ " & vbNewLine & _
             "From ҩƷ�շ���¼ a, �շ���ĿĿ¼ b, �շ���Ŀ���� c, ҩƷ��� d, ������ü�¼ e, ����ҽ����¼ f, ������ĿĿ¼ g, ҩƷ������ h " & vbNewLine & _
             "  ,Table(Cast(f_Str2list2([2], '|', ',') As t_Strlist2)) I "
    
    '��ͷ������ҽԺҪ��סԺ������;���ҩƷ����
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " , ҩƷ���� T "
    End If
    
    strSQL = strSQL & " Where a.ҩƷid = b.Id And a.ҩƷid = c.�շ�ϸĿid(+) And a.ҩƷid = d.ҩƷid And a.����id = e.Id and d.ҩ��id=g.id " & vbNewLine & _
             "    And e.ҽ����� = f.Id(+) And c.����(+) = 3 And a.����=h.����(+) And a.���� = i.c1 And a.NO = i.c2 " & _
             IIf(Trim(strStockIDs) = "", "", " And Instr([1], ';' || a.�ⷿid || ';') > 0 ")
    
    '��ͷ������ҽԺҪ��סԺ������;���ҩƷ����
    If gstrUnit = GCST_UNIT_BTSZXYY Then
        strSQL = strSQL & " And d.ҩ��id = t.ҩ��id And (a.���� = 9 And t.������� In ('����ҩ', '����I��') Or a.���� = 8) "
    End If
    
    mstrSQL = strSQL & vbCrLf & " union all  " & vbCrLf & Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
    
    If gintMode = 0 Then
        Set rsDetails = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    Else
        Set rsDetails = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeDetail_CQFLQZYY", ";" & strStockIDs & ";", strNO)
    End If
    
    strOutPutExeStep = "ִ�д�����ϢSQL�ɹ�"
    
    strXML = ""
    arrXML = Array()
    
    strOutPutExeStep = "ƴװXML��begin"
    
    With rsTemp
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_NO", SpecialChar(!�������))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!��ҩҩ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_ID", NVL(!���￨��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_NAME", SpecialChar(!��������))
                
                strOutPutExeStep = "ƴװ������XML��1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("INVOICE_NO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("PATIENT_TYPE", NVL(!��������))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DATE_OF_BIRTH", Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("SEX", SpecialChar(!�����Ա�))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_IDENTITY", SpecialChar(!�������))
                
                strOutPutExeStep = "ƴװ������XML��2"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("CHARGE_TYPE", SpecialChar(!ҽ������))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_ATTR", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_INFO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("RCPT_INFO", GetRCPT_INFO(NVL(!�������)))
                strDrug = strDrug & vbCrLf & GetXMLFormat("RCPT_REMARK", "")
                
                strOutPutExeStep = "ƴװ������XML��3"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("REPETITION", "1")
                strDrug = strDrug & vbCrLf & GetXMLFormat("COSTS", NVL(!����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PAYMENTS", NVL(!ʵ������))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ORDERED_BY", NVL(!�������ұ���))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ORDERED_BY_NAME", NVL(!������������))
                
                strOutPutExeStep = "ƴװ������XML��4"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESCRIBED_BY", SpecialChar(!����ҽ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("ENTERED_BY", SpecialChar(!¼����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSE_PRI", NVL(!��ҩ���ȼ�))
                
                strOutPutExeStep = "ƴװ������XML��5"
                
                '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
                rsDetails.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��)
                rsDetails.Sort = "���"
                
                strOutPutExeStep = "������ϸ��¼"
                
                strDetail = ""
                Do While Not rsDetails.EOF
                    strDetail = strDetail & vbCrLf & "<CONSIS_PRESC_DTLVW>"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PRESC_NO", NVL(rsDetails!no))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ITEM_NO", NVL(rsDetails!���))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ADVICE_CODE", NVL(rsDetails!ҽ�����))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(rsDetails!ҩƷ����))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(rsDetails!ҩƷ���))
                    
                    strOutPutExeStep = "ƴװ������ϸXML��1"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(rsDetails!ҩƷ����))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(rsDetails!�������ұ���))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(rsDetails!������������))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_SPEC", SpecialChar(rsDetails!ҩƷ��װ���))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_UNITS", SpecialChar(rsDetails!ҩƷ��λ))
                    
                    strOutPutExeStep = "ƴװ������ϸXML��2"
                     
                    strDetail = strDetail & vbCrLf & GetXMLFormat("QUANTITY", NVL(rsDetails!����))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("UNIT", SpecialChar(rsDetails!ҩƷ��λ))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("COSTS", NVL(rsDetails!����))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("PAYMENTS", NVL(rsDetails!ʵ������))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE", NVL(rsDetails!ҩƷ����))
                    
                    strOutPutExeStep = "ƴװ������ϸXML��3"
                    
                    strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE_UNITS", SpecialChar(rsDetails!������λ))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("ADMINISTRATION", SpecialChar(rsDetails!�÷�))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("FREQUENCY ", SpecialChar(rsDetails!ִ��Ƶ��))
                    strDetail = strDetail & vbCrLf & GetXMLFormat("Additionusage", "")
                    strDetail = strDetail & vbCrLf & GetXMLFormat("Rcpt_remark", "")
                    
                    strOutPutExeStep = "ƴװ������ϸXML��4"
                     
                    strDetail = strDetail & vbCrLf & "</CONSIS_PRESC_DTLVW>"
                    rsDetails.MoveNext
                Loop
                strDrug = strDrug & strDetail
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF And strXML <> "" Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    
                    strOutPutExeStep = "ƴװXML��end"
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeDetail_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeDetail_CQFLQZYY"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeDetail_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Private Function GetXML_RecipeDetailEx(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
'���ܣ�����ⷿIDΪ0��������ִ���ⷿID�벡��ID����XML�ַ���
'������
'  rsBill���������ݼ���
'  rsDetail����ϸ���ݼ���
'  varXML�����ɵ�XML�ַ������飨ʵ�Σ���
'���أ�True�ɹ�   Falseʧ��
'���ýӿڣ�Τ�ֺ���CONSISϵͳv2.2
    Const STR_ROOT_BEGIN = "<ROOT>"
    Const STR_ROOT_END = "</ROOT>"
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng�ⷿID As Long, lng����ID As Long
    Dim varReturn As Variant
    
    On Error GoTo errHandle
    varReturn = Array()
    With rsBill
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        lng�ⷿID = NVL(!��ҩҩ��, 0)
        lng����ID = NVL(!���￨��, 0)
        Do
            If .EOF Then Exit Do
            '����
            strBill = "<" & STR_BILL & " "
            strBill = strBill & vbCrLf & "PRESC_DATE = """ & Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "PRESC_NO = """ & SpecialChar(!�������) & """"
            strBill = strBill & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
            strBill = strBill & vbCrLf & "PATIENT_ID = """ & NVL(!���￨��) & """"
            strBill = strBill & vbCrLf & "PATIENT_NAME = """ & SpecialChar(!��������) & """"
            strBill = strBill & vbCrLf & "PATIENT_TYPE = """ & NVL(!��������) & """"
            strBill = strBill & vbCrLf & "DATE_OF_BIRTH = """ & Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss") & """"
            strBill = strBill & vbCrLf & "SEX = """ & SpecialChar(!�����Ա�) & """"
            strBill = strBill & vbCrLf & "PRESC_IDENTITY = """ & SpecialChar(!�������) & """"
            strBill = strBill & vbCrLf & "CHARGE_TYPE = """ & SpecialChar(!ҽ������) & """"
            strBill = strBill & vbCrLf & "PRESC_ATTR = """""
            strBill = strBill & vbCrLf & "PRESC_INFO = """""
            strBill = strBill & vbCrLf & "RCPT_INFO = " & GetRCPT_INFO(NVL(!�������))
            strBill = strBill & vbCrLf & "RCPT_REMARK = """""
            strBill = strBill & vbCrLf & "REPETITION = ""1"""
            strBill = strBill & vbCrLf & "COSTS = """ & NVL(!����) & """"
            strBill = strBill & vbCrLf & "PAYMENTS = """ & NVL(!ʵ������) & """"
            strBill = strBill & vbCrLf & "ORDERED_BY = """ & NVL(!��������) & """"
            strBill = strBill & vbCrLf & "PRESCRIBED_BY = """ & SpecialChar(!����ҽ��) & """"
            strBill = strBill & vbCrLf & "ENTERED_BY = """ & SpecialChar(!¼����) & """"
            strBill = strBill & vbCrLf & "DISPENSE_PRI = """ & NVL(!��ҩ���ȼ�) & """"
            strBill = strBill & vbCrLf & ">"
            
            '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
            strDetail = ""
            rsDetail.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��) & " and ����id=" & NVL(!���￨��)
            rsDetail.Sort = "���"
            Do
                If rsDetail.EOF Then Exit Do
                '��ϸ
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & " "
                strDetail = strDetail & vbCrLf & "PRESC_DATE = """ & Format(rsDetail!��������, "yyyy-MM-DDThh:mm:ss") & """"
                strDetail = strDetail & vbCrLf & "PRESC_NO = """ & NVL(rsDetail!no) & """"
                strDetail = strDetail & vbCrLf & "ITEM_NO = """ & rsDetail!��� & """"
                strDetail = strDetail & vbCrLf & "DRUG_CODE = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DRUG_NAME = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "TRADE_NAME = """ & SpecialChar(rsDetail!ҩƷ��Ʒ��) & """"
                strDetail = strDetail & vbCrLf & "DRUG_SPEC= """ & SpecialChar(rsDetail!ҩƷ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PACKAGE = """ & SpecialChar(rsDetail!ҩƷ��װ���) & """"
                strDetail = strDetail & vbCrLf & "DRUG_UNIT = """ & SpecialChar(rsDetail!ҩƷ��λ) & """"
                strDetail = strDetail & vbCrLf & "FIRM_ID = """ & SpecialChar(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DRUG_PRICE = """ & NVL(rsDetail!ҩƷ�۸�) & """"
                strDetail = strDetail & vbCrLf & "QUANTITY = """ & NVL(rsDetail!����) & """"
                strDetail = strDetail & vbCrLf & "COSTS = """ & NVL(rsDetail!����) & """"
                strDetail = strDetail & vbCrLf & "PAYMENTS = """ & NVL(rsDetail!ʵ������) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE = """ & NVL(rsDetail!ҩƷ����) & """"
                strDetail = strDetail & vbCrLf & "DOSAGE_UNITS = """ & SpecialChar(rsDetail!������λ) & """"
                strDetail = strDetail & vbCrLf & "ADMINISTRATION = """ & SpecialChar(rsDetail!�÷�) & """"
                strDetail = strDetail & vbCrLf & "FREQUENCY = """ & SpecialChar(rsDetail!ִ��Ƶ��) & """"
                strDetail = strDetail & vbCrLf & ">"
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '��ֲ�ͬ�ⷿID�Ͳ���ID�ĵ�����ϸ
            If lng�ⷿID = NVL(!��ҩҩ��, 0) And lng����ID = NVL(!���￨��, 0) Then
                strXML = strXML & strBill & vbCrLf
            Else
                strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
Call OutputLog(strXML)
                strXML = strBill & vbCrLf
            End If
            
            lng�ⷿID = NVL(!��ҩҩ��, 0)
            lng����ID = NVL(!���￨��, 0)
            
            .MoveNext
        Loop While Not .EOF
        
        strXML = STR_ROOT_BEGIN & vbCrLf & strXML & STR_ROOT_END
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        GetXML_RecipeDetailEx = True
Call OutputLog(strXML)

    End With
    
    Exit Function
    
errHandle:
    Set varXML = Nothing
End Function

'Public Sub OutPutLog(ByVal strOutput As String)
'    '���ڱ������û��������ԣ����������Ի򲻷����������ʱʹ��
'    '������ִ�еĹؼ����̣�����������ⲿ��־�ļ����Դ˷����������
'    'ע�⣺����Ҫ����ʱ�ֹ�����ָ������־�ļ������뻷��ʱ�ŵ�����̨��������Ŀ¼��Դ���뻷��ʱ�ŵ������ļ�����Ŀ¼
'    'ע�⣺�������Ҫ������Ҫ��ʱɾ����־�ļ���������־�ļ����ܻ��������ر����û������������������Ͽ�
'    '��ϵͳ����ָ����ͬ����־�ļ���
'    '��־�����Զ��壬�ο���ʽ��ʱ��+�����ڲ�����/����+ҵ������/����+�ؼ�����
'    'Ĭ�ϵĴ�������ʱ�䣬�������Ҫ����ȥ��
'    Dim objFile As New FileSystemObject
'    Dim objTarget As TextStream
'    Const STR_CONS_FILENAME As String = "zlDrugPacker.log"
'
'    Err = 0
'
'    On Error Resume Next
'
'    '����ļ��Ƿ����
'    Set objTarget = objFile.OpenTextFile(App.Path & "\" & STR_CONS_FILENAME)
'
'    '������������������
'    If objTarget Is Nothing Then Exit Sub
'
''    If err <> 0 Then
''        '����Ŀ���ļ�
''        Set objFile = CreateObject("Scripting.FileSystemObject")
''        Set objTarget = objFile.CreateTextFile(App.Path & "\" & STR_CONS_FILENAME, True)
''        objTarget.Close
''    End If
'
'    Err.Clear
'    On Error GoTo errHand
'
'    Open App.Path & "\" & STR_CONS_FILENAME For Append Shared As #1
'
'    Print #1, strOutput
'    Close #1
'
'    Exit Sub
'errHand:
'    Close #1
''    MsgBox err.Description, vbExclamation + vbOKOnly
'End Sub

Public Sub OutputLog(ByVal strOutput As String)
'���ܣ�����������д���ض����ռ��ļ���
'������
'  strOutput���ռ�����

    Const STR_LOG_FILENAME As String = "zlDrugPacker"       '��־�ı�����
    Const INT_MAX_DAY As Integer = 7                        '��־��������

    Dim objTS As TextStream
    Dim objFolder As Folder
    Dim objFile As File
    Dim strDate As String, strFileName As String
    Dim blnExist As Boolean, blnAutoCreate As Boolean

    On Error GoTo hErr

    '��ȡע���Ĳ���
    blnAutoCreate = Val(GetSetting("ZLSOFT", "����ģ��\�Զ���ҩ��", "�Զ�������־")) = 1

    If blnAutoCreate Then
        '�Զ�������־�ļ�
        
        strFileName = STR_LOG_FILENAME & Format(Date, "_yyyymmdd") & ".log"
    
        ''�ж��ļ��Ƿ����
        Set objFolder = mobjFSO.GetFolder(App.Path)
        For Each objFile In objFolder.Files
            If LCase(objFile.Name) Like LCase(strFileName) Then
                blnExist = True
                Exit For
            End If
        Next
        
        Set objTS = mobjFSO.OpenTextFile(App.Path & "\" & strFileName, ForAppending, True)
        If blnExist = False Then
            '�´������ļ���ǿ�Ƽ���ʱ���
            strOutput = Now() & vbCrLf & strOutput
        End If
        objTS.WriteLine strOutput
        objTS.Close
        
        ''������������־�ļ�����ɾ��
        Set objFolder = mobjFSO.GetFolder(App.Path)
        For Each objFile In objFolder.Files
            If LCase(objFile.Name) Like LCase(STR_LOG_FILENAME) & "_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].log" Then
                strDate = Split(objFile.Name, "_")(1)
                strDate = Split(strDate, ".")(0)
                strDate = Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2)
                If Abs(Date - CDate(strDate)) >= INT_MAX_DAY Then
                    On Error Resume Next
                    objFile.Delete True
                    On Error GoTo hErr
                End If
            End If
        Next
    
    Else
        '���Զ�������־�ļ�
        
        strFileName = STR_LOG_FILENAME & ".log"
    
        ''��Ӵ洢��־��ʽ
        Set objTS = mobjFSO.OpenTextFile(App.Path & "\" & strFileName, ForAppending)
        If objTS Is Nothing Then Exit Sub
        
        objTS.WriteLine strOutput
        objTS.Close
    End If
    
    Exit Sub
    
hErr:
End Sub

Private Function GetXML_RecipeDetailEx_CQFLQZYY(ByVal rsBill As ADODB.Recordset, ByVal rsDetail As ADODB.Recordset, ByRef varXML As Variant) As Boolean
    '���ܣ�����ⷿIDΪ0��������ִ���ⷿID�벡��ID����XML�ַ���
    '������
    '  rsBill���������ݼ���
    '  rsDetail����ϸ���ݼ���
    '  varXML�����ɵ�XML�ַ������飨ʵ�Σ���
    '���أ�True�ɹ�   Falseʧ��
    '���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Const STR_BILL = "CONSIS_PRESC_MSTVW"
    Const STR_DETAIL = "CONSIS_PRESC_DTLVW"
    Dim strXML As String, strBill As String, strDetail As String
    Dim lng�ⷿID As Long, lng����ID As Long
    Dim varReturn As Variant
    Dim strOutput As String           '������־���
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
    strOutput = "���ú�����GetXML_RecipeDetailEx_CQFLQZYY"
    
    On Error GoTo errHandle
    
    varReturn = Array()
    
    strOutPutExeStep = "��ʼ������"
    
    With rsBill
        If .RecordCount <= 0 Then
            strOutput = strOutput & vbCrLf & "������"
            strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeDetailEx_CQFLQZYY"
            Call OutputLog(strOutput)
            
            Exit Function
        End If
       
        .MoveFirst
        
        strOutPutExeStep = "�ⷿID������ID��ʼ��ֵ"
        
        lng�ⷿID = NVL(!��ҩҩ��, 0)
        lng����ID = NVL(!���￨��, 0)
        
        strOutPutExeStep = "ƴװXML��begin"
        
        Do
            If .EOF Then Exit Do
            '����
            strBill = "<" & STR_BILL & ">"
            
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!����ʱ��, "yyyy-MM-DDThh:mm:ss"))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_NO", SpecialChar(!�������))
            strBill = strBill & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!��ҩҩ��))
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_ID", NVL(!���￨��))
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_NAME", SpecialChar(!��������))
            
            strOutPutExeStep = "ƴװ������XML��1"
            
            strBill = strBill & vbCrLf & GetXMLFormat("INVOICE_NO", "")
            strBill = strBill & vbCrLf & GetXMLFormat("PATIENT_TYPE", NVL(!��������))
            strBill = strBill & vbCrLf & GetXMLFormat("DATE_OF_BIRTH", Format(NVL(!���߳�������), "yyyy-MM-DDThh:mm:ss"))
            strBill = strBill & vbCrLf & GetXMLFormat("SEX", SpecialChar(!�����Ա�))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_IDENTITY", SpecialChar(!�������))
            
            strOutPutExeStep = "ƴװ������XML��2"
            
            strBill = strBill & vbCrLf & GetXMLFormat("CHARGE_TYPE", SpecialChar(!ҽ������))
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_ATTR", "")
            strBill = strBill & vbCrLf & GetXMLFormat("PRESC_INFO", "")
            strBill = strBill & vbCrLf & GetXMLFormat("RCPT_INFO", GetRCPT_INFO(NVL(!�������)))
            strBill = strBill & vbCrLf & GetXMLFormat("RCPT_REMARK", "")
            
            strOutPutExeStep = "ƴװ������XML��3"
            
            strBill = strBill & vbCrLf & GetXMLFormat("REPETITION", "1")
            strBill = strBill & vbCrLf & GetXMLFormat("COSTS", NVL(!����))
            strBill = strBill & vbCrLf & GetXMLFormat("PAYMENTS", NVL(!ʵ������))
            strBill = strBill & vbCrLf & GetXMLFormat("ORDERED_BY", NVL(!�������ұ���))
            strBill = strBill & vbCrLf & GetXMLFormat("ORDERED_BY_NAME", NVL(!������������))
            
            strOutPutExeStep = "ƴװ������XML��4"
            
            strBill = strBill & vbCrLf & GetXMLFormat("PRESCRIBED_BY", SpecialChar(!����ҽ��))
            strBill = strBill & vbCrLf & GetXMLFormat("ENTERED_BY", SpecialChar(!¼����))
            strBill = strBill & vbCrLf & GetXMLFormat("DISPENSE_PRI", NVL(!��ҩ���ȼ�))
            
            strOutPutExeStep = "ƴװ������XML��5"
            
            '������ϸ��¼��ȷ���뵥�ݶ�Ӧ
            strDetail = ""
            rsDetail.Filter = "no='" & !������� & "' and ����=" & NVL(!����) & " and �ⷿid=" & NVL(!��ҩҩ��) & " and ����id=" & NVL(!���￨��)
            rsDetail.Sort = "���"
            
            strOutPutExeStep = "������ϸ��¼"
            
            Do
                If rsDetail.EOF Then Exit Do
                '��ϸ
                strDetail = strDetail & vbCrLf & "<" & STR_DETAIL & ">"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("PRESC_NO", NVL(rsDetail!no))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ITEM_NO", NVL(rsDetail!���))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ADVICE_CODE", NVL(rsDetail!ҽ�����))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(rsDetail!ҩƷ����))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_SPEC", SpecialChar(rsDetail!ҩƷ���))
                
                strOutPutExeStep = "ƴװ������ϸXML��1"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("DRUG_NAME", SpecialChar(rsDetail!ҩƷ����))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_ID", SpecialChar(rsDetail!�������ұ���))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FIRM_NAME", SpecialChar(rsDetail!������������))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_SPEC", SpecialChar(rsDetail!ҩƷ��װ���))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PACKAGE_UNITS", SpecialChar(rsDetail!ҩƷ��λ))
                
                strOutPutExeStep = "ƴװ������ϸXML��2"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("QUANTITY", NVL(rsDetail!����))
                strDetail = strDetail & vbCrLf & GetXMLFormat("UNIT", SpecialChar(rsDetail!ҩƷ��λ))
                strDetail = strDetail & vbCrLf & GetXMLFormat("COSTS", NVL(rsDetail!����))
                strDetail = strDetail & vbCrLf & GetXMLFormat("PAYMENTS", NVL(rsDetail!ʵ������))
                strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE", NVL(rsDetail!ҩƷ����))
                
                strOutPutExeStep = "ƴװ������ϸXML��3"
                
                strDetail = strDetail & vbCrLf & GetXMLFormat("DOSAGE_UNITS", SpecialChar(rsDetail!������λ))
                strDetail = strDetail & vbCrLf & GetXMLFormat("ADMINISTRATION", SpecialChar(rsDetail!�÷�))
                strDetail = strDetail & vbCrLf & GetXMLFormat("FREQUENCY ", SpecialChar(rsDetail!ִ��Ƶ��))
                strDetail = strDetail & vbCrLf & GetXMLFormat("Additionusage", "")
                strDetail = strDetail & vbCrLf & GetXMLFormat("Rcpt_remark", "")
                
                strOutPutExeStep = "ƴװ������ϸXML��4"
                
                strDetail = strDetail & vbCrLf & "</" & STR_DETAIL & ">"
                rsDetail.MoveNext
            Loop While Not rsDetail.EOF
            
            strBill = strBill & strDetail
            strBill = strBill & "</" & STR_BILL & ">"
            
            '��ֲ�ͬ�ⷿID�Ͳ���ID�ĵ�����ϸ
            If lng�ⷿID = NVL(!��ҩҩ��, 0) And lng����ID = NVL(!���￨��, 0) Then
                strXML = strXML & strBill & vbCrLf
                strOutPutExeStep = "��ֲ�ͬ�ⷿID�Ͳ���ID�ĵ�����ϸ"
            Else
                ReDim Preserve varReturn(UBound(varReturn) + 1)
                varReturn(UBound(varReturn)) = strXML
                strXML = strBill & vbCrLf
                
                strOutPutExeStep = "����ͬ�ⷿ�Ͳ��˷ֱ�װ��������"
            End If
            
            lng�ⷿID = NVL(!��ҩҩ��, 0)
            lng����ID = NVL(!���￨��, 0)
            
            strOutPutExeStep = "�ⷿID������ID��ֵ"
            
            .MoveNext
        Loop While Not .EOF
        
        ReDim Preserve varReturn(UBound(varReturn) + 1)
        varReturn(UBound(varReturn)) = strXML
        varXML = varReturn
        
        strOutPutExeStep = "ƴװXML��end"
    End With
    
    GetXML_RecipeDetailEx_CQFLQZYY = True
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeDetailEx_CQFLQZYY"
    Call OutputLog(strOutput)
            
    Exit Function
    
errHandle:
    Set varXML = Nothing
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeDetail_CQFLQZYY"
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    Call OutputLog(strOutput)
End Function


Public Function GetXML_RecipeList(ByVal lngStockID As Long, ByVal strNO As String) As Variant
'����������֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mstrSQL = "Select ��������,No From ҩƷ�շ���¼ Where �ⷿid=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mstrSQL = mstrSQL & " And ����=[2] And NO=[3]"
    Else
        mstrSQL = mstrSQL & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mstrSQL = mstrSQL & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mstrSQL = mstrSQL & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mstrSQL = mstrSQL & ")"
    End If
    mstrSQL = mstrSQL & " and (��¼״̬=1 or mod(��¼״̬,3)=1) "
    
    If InStr(1, strNO, "|") < 1 Then
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        End If
    Else
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID)
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList", lngStockID)
        End If
    End If
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW"
                strDrug = strDrug & vbCrLf & "PRESC_DATE = """ & Format(!��������, "yyyy-MM-DDThh:mm:ss") & """"
                strDrug = strDrug & vbCrLf & "PRESC_NO = """ & NVL(!no) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                    Call OutputLog(strXML)
                End If
            Loop
        End If
    End With
    
    GetXML_RecipeList = arrXML
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_RecipeList_CQFLQZYY(ByVal lngStockID As Long, ByVal strNO As String, ByRef strOutput As String) As Variant
    '����������֯��ָ����XML��ʽ
    '���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant, arrTmp As Variant
    Dim i As Integer
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
'    Presc_date  ����ʱ��    Datetime    N
'    Presc_no    �������    Nvarchar(200)   N
'    Invoice_no  ��Ʊ���    Nvarchar(200)   Y
'    DISPENSARY  ��ҩҩ�ֱ��
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_RecipeList_CQFLQZYY"
   
    mstrSQL = "Select ��������,No From δ��ҩƷ��¼ Where �ⷿid=[1]"
    
    If InStr(1, strNO, "|") < 1 Then
        mstrSQL = mstrSQL & " And ����=[2] And NO=[3]"
    Else
        mstrSQL = mstrSQL & " And ("
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(arrTmp)
            If i = UBound(arrTmp) Then
                mstrSQL = mstrSQL & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "')"
            Else
                mstrSQL = mstrSQL & "(����=" & Split(arrTmp(i), ",")(0) & " And NO='" & Split(arrTmp(i), ",")(1) & "') or "
            End If
        Next
        mstrSQL = mstrSQL & ")"
    End If
    
    strOutPutExeStep = "ƴ��SQL"
    
    If InStr(1, strNO, "|") < 1 Then
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
        End If
    Else
        If gintMode = 0 Then
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID)
        Else
            Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_RecipeList_CQFLZYY", lngStockID)
        End If
    End If
    
    strOutPutExeStep = "ִ��SQL�ɹ�"
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        strOutPutExeStep = "ƴװXML��begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_PRESC_MSTVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_DATE", Format(!��������, "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRESC_NO", NVL(!no))
                strDrug = strDrug & vbCrLf & GetXMLFormat("INVOICE_NO", "")
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", lngStockID)
                
                strOutPutExeStep = "ƴװXML��1"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_PRESC_MSTVW>"
                
                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
            
            strOutPutExeStep = "ƴװXML��end"
        End If
    End With
    
    GetXML_RecipeList_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeList_CQFLQZYY"
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeList_CQFLQZYY"
    Call OutputLog(strOutput)
End Function

Public Function IsRegisterStock(ByVal lngStockID As Long, ByVal strStockIDs As String) As Boolean
'���ܣ�����Ƿ�Ϊע��ҩ��
    Dim i As Integer
    Dim arrID As Variant
    
    If Val(strStockIDs) = 0 Or lngStockID = 0 Then Exit Function
    
    arrID = Split(strStockIDs, ";")
    For i = LBound(arrID) To UBound(arrID)
        If Val(arrID(i)) = lngStockID Then
            IsRegisterStock = True
            Exit For
        End If
    Next
End Function

Public Function GetXML_RecipeReturn_CQFLQZYY(ByVal strReturnRecipt As String, ByVal strStockIDs As String, ByRef strOutput As String) As Variant
'����������֯��ָ����XML��ʽ
'�˷Ѵ�����Ϣ
'strReturnRecipt���˷Ѵ�����Ϣ����ʽ��NO,ҩ��id|NO,ҩ��id
'���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Dim strXML As String
    Dim arrXML As Variant
    Dim arrRecipt
    Dim n As Integer
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
'    Presc_date  ����ʱ��    Datetime    N
'    Presc_no    �������    Nvarchar(200)   N
'    DISPENSARY  ��ҩҩ�ֱ��

    strOutput = strOutput & vbCrLf & "���ú�����GetXML_RecipeReturn_CQFLQZYY"
    
    On Error GoTo errHandle
    
    arrRecipt = Split(strReturnRecipt, "|")
    arrXML = Array()
    
    strOutPutExeStep = "ƴװXML��begin"
    
    For n = 0 To UBound(arrRecipt)
        'ע���ҩ�����ύ����
        If IsRegisterStock(Val(Split(arrRecipt(n), ",")(1)), strStockIDs) Then
            strXML = IIf(strXML = "", "", strXML & vbCrLf) & "<CONSIS_PRESC_MSTVW>"
                        
            strXML = strXML & vbCrLf & GetXMLFormat("PRESC_DATE", Format(CStr(Now), "yyyy-MM-DDThh:mm:ss"))
            strXML = strXML & vbCrLf & GetXMLFormat("PRESC_NO", Split(arrRecipt(n), ",")(0))
            strXML = strXML & vbCrLf & GetXMLFormat("INVOICE_NO", "")
            strXML = strXML & vbCrLf & GetXMLFormat("DISPENSARY", Split(arrRecipt(n), ",")(1))
            
            strXML = strXML & vbCrLf & "</CONSIS_PRESC_MSTVW>"
        End If
    Next
    
    If strXML <> "" Then
        ReDim Preserve arrXML(UBound(arrXML) + 1)
        arrXML(UBound(arrXML)) = strXML
    End If
    
    strOutPutExeStep = "ƴװXML��end"
    
    GetXML_RecipeReturn_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_RecipeReturn_CQFLQZYY"
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_RecipeReturn_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXML_Stock(ByVal lngStockID As Long) As Variant
'��ҩƷ�����Ϣ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv2.2
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim strTitle As String
    Dim arrXML As Variant
    
    On Error GoTo errHandle
    mstrSQL = "Select a.id ҩƷ���,c.�ⷿid ��ҩҩ��,sum(c.ʵ������/e.�����װ) ҩƷ����,d.�ⷿ��λ ҩƷ��λ " & vbNewLine & _
              "From �շ���ĿĿ¼ a, ҩƷ��� c, ҩƷ�����޶� d,ҩƷ��� e " & vbNewLine & _
              "Where a.Id = c.ҩƷid And e.ҩƷid=c.ҩƷid And d.�ⷿid(+) = c.�ⷿid And d.ҩƷid(+) = c.ҩƷid And c.�ⷿid=[1] " & vbNewLine & _
              "Group By a.id, c.�ⷿid, d.�ⷿ��λ " & vbNewLine & _
              "Having Sum(c.ʵ������/e.�����װ)<>0 "
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Stock", lngStockID)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Stock", lngStockID)
    End If
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        If .RecordCount > 0 Then
            strTitle = "<ROOT>"
            
            Do While Not .EOF
                strDrug = "<CONSIS_PHC_STORAGEVW"
                strDrug = strDrug & vbCrLf & "DRUG_CODE = """ & SpecialChar(!ҩƷ���) & """"
                strDrug = strDrug & vbCrLf & "DISPENSARY = """ & NVL(!��ҩҩ��) & """"
                strDrug = strDrug & vbCrLf & "DRUG_QUANTITY = """ & NVL(!ҩƷ����) & """"
                strDrug = strDrug & vbCrLf & "LOCATIONINFO = """ & SpecialChar(!ҩƷ��λ) & """"
                strDrug = strDrug & vbCrLf & ">"
                strDrug = strDrug & vbCrLf & "</CONSIS_PHC_STORAGEVW>"

'��ҵ���ܿ��Բ���4K����
                strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
                
'                If Len(strXML & strDrug) > 3900 Then
'                    '����ǰ����ӵ�����
'                    strXML = strXML & vbCrLf & "</ROOT>"
'                    ReDim Preserve arrXML(UBound(arrXML) + 1)
'                    arrXML(UBound(arrXML)) = strXML
'
'                    '����ƴ���µ�XML
'                    strXML = strTitle & vbCrLf & strDrug
'                Else
'                    strXML = IIf(strXML = "", strTitle, strXML) & vbCrLf & strDrug
'                End If
                
                rsTemp.MoveNext
                If .EOF Then
                    strXML = strXML & vbCrLf & "</ROOT>"
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
        End If
    End With
    
    GetXML_Stock = arrXML
    Exit Function
    
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Dept(ByVal strProperty As String, Optional ByRef strLog As String) As Variant
'���ܣ���ȡZLHIS�Ĳ������ݣ���ת���ɡ�Τ�ֺ��ġ��ӿ�Ҫ���XML��ʽ
'������
'  strLog����־����
'���أ�XML�ַ�������

    Dim rsTemp As ADODB.Recordset
    Dim arrXML As Variant
    Dim objXML As clsXML
    
strLog = strLog & "��ȡ�������ݿ�ʼ��" & vbCrLf
    
    mstrSQL = "Select Distinct a.Id, a.����, b.������� " & vbNewLine & _
              "From ���ű� A, ��������˵�� B " & vbNewLine & _
              "Where a.Id = b.����id And Trunc(Nvl(a.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))) = To_Date('3000-1-1', 'YYYY-MM-DD') " & _
              "    And b.������� <> 0 "
    
    If strProperty <> "" Then
        strProperty = "," & strProperty & ","
        mstrSQL = mstrSQL & " And Instr([1], ',' || b.�������� || ',') > 0 "
    End If
    mstrSQL = mstrSQL & vbNewLine & _
              "Order By a.ID "
              
    On Error GoTo hErr
    
strLog = strLog & "SQL��" & mstrSQL & vbCrLf
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Dept", strProperty)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Dept", strProperty)
    End If
    
strLog = strLog & "׼����װXML��" & vbCrLf
    
    With rsTemp
        arrXML = Array()
        Set objXML = New clsXML
        Do While .EOF = False
            objXML.ClearXmlText
            Call objXML.AppendNode("CONSIS_BASIC_DEPTVW")
                Call objXML.AppendData("DEPTCODE", NVL(!ID))
                Call objXML.AppendData("DEPTNAME", NVL(!����))
                Call objXML.AppendData("OUTP_OR_INP", NVL(!�������))
            Call objXML.AppendNode("CONSIS_BASIC_DEPTVW", True)
            
'strLog = strLog & "XML��" & objXML.XmlText
            
            ReDim Preserve arrXML(UBound(arrXML) + 1)
            arrXML(UBound(arrXML)) = objXML.XmlText
            
'strLog = strLog & "��ɣ�" & vbCrLf
            
            .MoveNext
        Loop
        .Close
    End With
    
strLog = strLog & "��װXML��ɣ�" & vbCrLf
    
    GetXML_Dept = arrXML
    
    Exit Function
    
hErr:
strLog = strLog & "��ȡ���������쳣��" & Err.Description & vbCrLf
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
End Function

Public Function GetXML_Stock_CQFLQZYY(ByVal lngStockID As Long, ByRef strOutput As String) As Variant
'��ҩƷ�����Ϣ��֯��ָ����XML��ʽ
'���ýӿڣ�Τ�ֺ���CONSISϵͳv4.3
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim strDrug As String
    Dim arrXML As Variant
    Dim strOutPutExeStep As String    'ִ�в��裬���������־�����������
    
'    �ֶ���  ����    ˵��    NULL
'    Dispensary  Nvarchar(40)    ��ҩҩ��    N
'    Drug_code   Nvarchar(40)    ҩƷ���    N
'    Locationinfo    Nvarchar(200)   ��λ��Ϣ    N
'    Batchid Nvarchar(200)   ҩƷ����    Y
'    Batchno Nvarchar(200)   ҩƷ����    Y
'    Producedate Datetime    ��������    Y
'    Disableddate    Datetime    ʧЧ����    Y
'    Quantity    Numeric(20,6)   ҩƷ��λ�������    N
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "���ú�����GetXML_Stock_CQFLQZYY"

    mstrSQL = "Select a.���� ҩƷ���, c.�ⷿid ��ҩҩ��, c.ʵ������ / e.�����װ As ҩƷ����, Nvl(d.�ⷿ��λ, '��') As ҩƷ��λ, Nvl(c.����, 0) As ����, c.�ϴ�����, c.Ч��, c.�ϴ���������" & vbNewLine & _
        " From �շ���ĿĿ¼ A, ҩƷ��� C, ҩƷ�����޶� D, ҩƷ��� E" & vbNewLine & _
        " Where a.Id = c.ҩƷid And e.ҩƷid = c.ҩƷid And d.�ⷿid(+) = c.�ⷿid And d.ҩƷid(+) = c.ҩƷid And c.�ⷿid = [1] " & vbNewLine & _
        " Order By a.Id "
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "GetXML_Stock_CQFLQZYY", lngStockID)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "GetXML_Stock_CQFLQZYY", lngStockID)
    End If
    
    strOutPutExeStep = "ִ��SQL�ɹ�"
    
    strXML = ""
    arrXML = Array()
    
    With rsTemp
        strOutPutExeStep = "ƴװXML��begin"
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                strDrug = "<CONSIS_BASIC_LOCATIONVW>"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISPENSARY", NVL(!��ҩҩ��))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_CODE", SpecialChar(!ҩƷ���))
                strDrug = strDrug & vbCrLf & GetXMLFormat("LOCATIONINFO", SpecialChar(!ҩƷ��λ))
                strDrug = strDrug & vbCrLf & GetXMLFormat("BATCHID", NVL(!����))
                strDrug = strDrug & vbCrLf & GetXMLFormat("BATCHNO", SpecialChar(NVL(!�ϴ�����)))
                
                strOutPutExeStep = "ƴװXML��1"
                
                strDrug = strDrug & vbCrLf & GetXMLFormat("PRODUCEDATE", Format(NVL(!�ϴ���������), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DISABLEDDATE", Format(NVL(!Ч��), "yyyy-MM-DDThh:mm:ss"))
                strDrug = strDrug & vbCrLf & GetXMLFormat("DRUG_QUANTITY", NVL(!ҩƷ����))
                
                strOutPutExeStep = "ƴװXML��2"
                
                strDrug = strDrug & vbCrLf & "</CONSIS_BASIC_LOCATIONVW>"

                strXML = IIf(strXML = "", "", strXML & vbCrLf) & strDrug
                
                rsTemp.MoveNext
                
                If .EOF Then
                    ReDim Preserve arrXML(UBound(arrXML) + 1)
                    arrXML(UBound(arrXML)) = strXML
                End If
            Loop
            
            strOutPutExeStep = "ƴװXML��end"
        End If
    End With
    
    GetXML_Stock_CQFLQZYY = arrXML
    
    strOutput = strOutput & vbCrLf & "ִ�гɹ���GetXML_Stock_CQFLQZYY"
 
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If

    strOutput = strOutput & vbCrLf & "�����쳣����"
    strOutput = strOutput & vbCrLf & "����裺" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "���SQL" & vbCrLf & mstrSQL
    strOutput = strOutput & vbCrLf & "ִ��ʧ�ܣ�GetXML_Stock_CQFLQZYY"
    Call OutputLog(strOutput)
End Function


Public Function GetXMLFormat(ByVal strNode As String, ByVal strText As String, Optional ByVal blnNodeUpper As Boolean = True) As String
    '����ڵ�����ݣ���ϳ�XML��ʽ
    '��ʽ��<NODE>Text</NODE>
    strNode = Replace(strNode, "<", "")
    strNode = Replace(strNode, "</", "")
    strNode = Replace(strNode, ">", "")
    If blnNodeUpper = True Then
        GetXMLFormat = "<" & UCase(strNode) & ">" & strText & "</" & UCase(strNode) & ">"
    Else
        GetXMLFormat = "<" & strNode & ">" & strText & "</" & strNode & ">"
    End If
End Function

Public Function SetSendWin(ByVal lngStockID As Long, ByVal strNO As String, ByVal intOpr As Integer) As Boolean
'����HIS��ָ�������ķ�ҩ����
    Dim i As Integer
    Dim arrTmp As Variant
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mstrSQL = "Select ���� From ��ҩ���� Where ҩ��id=[1] And ����=[2]"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, "SetSendWin", lngStockID, CStr(intOpr))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(mstrSQL, "SetSendWin", lngStockID, CStr(intOpr))
    End If
    
    If Not rsTemp.EOF Then
        arrTmp = Split(strNO, "|")
        For i = 0 To UBound(Split(strNO, "|"))
            mstrSQL = "Zl_δ��ҩƷ��¼_���䷢ҩ����("
            mstrSQL = mstrSQL & "'" & Split(arrTmp(i), ",")(1) & "',"
            mstrSQL = mstrSQL & Split(arrTmp(i), ",")(0) & ","
            mstrSQL = mstrSQL & lngStockID & ","
            mstrSQL = mstrSQL & "'" & rsTemp!���� & "')"
            
            Call OutputLog(mstrSQL)
            If gintMode = 0 Then
                Call gobjComLib.zlDatabase.ExecuteProcedure(mstrSQL, "SetSendWin")
            Else
                Call mdlDrugPacker.ExecuteProcedure(mstrSQL, "SetSendWin")
            End If
        Next
        SetSendWin = True
    Else
        If gblnShowMsg Then
            MsgBox "û���ҵ�����Ϊ��" & intOpr & "���Ĵ��ڣ����顰��ҩ���ڹ���ģ�飡", vbCritical, GSTR_MESSAGE
        Else
            Call OutputLog("û���ҵ�����Ϊ��" & intOpr & "���Ĵ��ڣ����顰��ҩ���ڹ���ģ�飡")
        End If
    End If
    
    Exit Function
    
errHandle:
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter() = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    Call OutputLog("SetSendWin�쳣�� " & Err.Description)
End Function


Public Function GetLocalIP() As String
'ȡ����IP
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    
    
    On Error GoTo EndRow
        GetIpAddrTable ByVal 0&, Ret, True
    
    
        If Ret <= 0 Then Exit Function
        ReDim bBytes(0 To Ret - 1) As Byte
        ReDim TempList(0 To Ret - 1) As String
        
        'retrieve the data
        GetIpAddrTable bBytes(0), Ret, False
          
        'Get the first 4 bytes to get the entry's.. ip installed
        CopyMemory Listing.dEntrys, bBytes(0), 4
        
        For Tel = 0 To Listing.dEntrys - 1
            'Copy whole structure to Listing..
            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Next Tel
        'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        GetLocalIP = TempIP 'Return The TempIP
    Exit Function
EndRow:
    GetLocalIP = ""
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Private Function GetRCPT_INFO(ByVal strNO As String) As String
'���ܣ���ȡ�����Ϣ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select MAX(DECODE(Id,1,�������,''))||';'||MAX(DECODE(Id,2,�������,'')) as ��� " & vbNewLine & _
             "From ( " & vbNewLine & _
             "      Select Rownum As Id,������� " & vbNewLine & _
             "      From (Select �������||decode(�Ƿ�����,1,'?','') ������� " & vbNewLine & _
             "            From ������ϼ�¼ " & vbNewLine & _
             "            Where ����id=(Select distinct ����id " & vbNewLine & _
             "                          From ( Select a.����id From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ�����=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And ��¼����=1 ) ) " & vbNewLine & _
             "              And ��ҳid=(Select distinct Case When ��ҳid Is Null Then (Select Id From ���˹Һż�¼ Where No=c.�Һŵ�) Else ��ҳId End As ��ҳid " & vbNewLine & _
             "                          From ( Select null ��ҳid, b.�Һŵ� From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ�����=b.Id " & vbNewLine & _
             "                                 Where a.No=[1] And ��¼����=1 ) c ) " & vbNewLine & _
             "union all " & vbNewLine & _
             "Select a.ժҪ As ������� From ���˹Һż�¼ a " & vbNewLine & _
             "Where No= (Select distinct Case When b.�Һŵ� Is Null Then ' ' Else b.�Һŵ� End As No " & vbNewLine & _
             "           From ������ü�¼ a Left Join ����ҽ����¼ b On a.ҽ����� = b.Id " & vbNewLine & _
             "           Where a.No = [1] And ��¼���� = 1 ) ) ) "
    On Error GoTo errHandle
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����Ϣ", strNO)
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSQL, "��ȡ�����Ϣ", strNO)
    End If
    
    If Not rsTemp.EOF Then
        GetRCPT_INFO = IIf(Trim(NVL(rsTemp!���)) = ";", """""", """" & Trim(NVL(rsTemp!���)) & """")
    Else
        GetRCPT_INFO = """"""
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    GetRCPT_INFO = """"""
End Function

Private Function SpecialChar(ByVal strVal As Variant) As String
'���ܣ������ַ�ת��
'˵����
' < ת &lt;
' > ת &gt;
' & ת &amp;
' ' ת &apos;
' " ת &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Private Function GetStockID(ByVal strText As String) As Long
    '���ܣ���ȡXML�ı��е�ҩ��ID
    Const STR_KEY = "DISPENSARY = "
    Const STR_NODES_BEGIN = "<DISPENSARY>"
    Const STR_NODES_END = "</DISPENSARY>"
    
    Dim lngStockID As Long
    Dim intStart As Integer
    Dim strTmp As String
    
    If strText = "" Then Exit Function
    
    intStart = InStr(strText, STR_KEY)
    If intStart > 0 Then
        lngStockID = Val(Mid(strText, intStart + Len(STR_KEY) + 1))
    Else
        strTmp = Mid(strText, 1, InStr(1, strText, STR_NODES_END) - 1)
        lngStockID = Val(Mid(strTmp, InStr(1, strTmp, STR_NODES_BEGIN) + Len(STR_NODES_BEGIN)))
    End If
    
    GetStockID = lngStockID
    
End Function

Public Function PackingWindow_DYEY(ByVal strNO As String, Optional ByRef strOut As String) As String
'���ܣ���ȡ���뵥�ݺŵĲ���ID����ҩҩ������ҩ���ڡ�ҩƷ������Ϣ����Ҫ���ƶ�ҵ����ʹ��
'������
'  strNO��������Ϣ����ʽ������ò�˵��
'  strOut��ʵ�Σ����쳣��Ϣ
'���أ�����ID����ҩҩ������ҩ���ڡ�ҩƷ������Ϣ
'XML��ʽ��
'<OUTPUT>
'  <BRID>����ID</BRID>
'  <ITEM>
'    <YFMC>ҩ������</YFMC>
'    <YFCK>��ҩ����</YFCK>
'    <YFMX>
'      <ITEM>
'        <MC>ҩƷ����1</MC>
'      </ITEM>
'      <ITEM>
'        <MC>ҩƷ����2</MC>
'      </ITEM>
'      <ITEM>
'        <MC>ҩƷ����...</MC>
'      </ITEM>
'    </YFMX>
'  </ITEM>
'  <ITEM>
'    ...
'  </ITEM>
'</OUTPUT>

    Const STR_OUT_B As String = "<OUTPUT>", STR_OUT_E As String = "</OUTPUT>"
    Const STR_BRID_B As String = "<BRID>", STR_BRID_E As String = "</BRID>"
    Const STR_ITEM_B As String = "<ITEM>", STR_ITEM_E As String = "</ITEM>"
    Const STR_YFMC_B As String = "<YFMC>", STR_YFMC_E As String = "</YFMC>"
    Const STR_YFCK_B As String = "<YFCK>", STR_YFCK_E As String = "</YFCK>"
    Const STR_YFMX_B As String = "<YFMX>", STR_YFMX_E As String = "</YFMX>"
    Const STR_MC_B As String = "<MC>", STR_MC_E As String = "</MC>"

    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String, strReturn As String, strDrugs As String
    Dim strStore As String, strWin As String
    Dim lngStoreID As Long
    
    On Error GoTo errHandle
    
    strSQL = "Select Distinct b.����id, a.�ⷿid, d.���� As ҩ������, a.��ҩ����, c.���� As ҩƷ���� " & vbCr & _
             "From ҩƷ�շ���¼ A, ������ü�¼ B, �շ���ĿĿ¼ C, ���ű� D, Table(f_Str2list2([1], '|', ',')) E " & vbCr & _
             "Where a.����id = b.Id And a.ҩƷid = c.Id And a.�ⷿid = d.Id And a.���� = e.C1 And a.No = e.C2 " & vbCr & _
             "Order By Nvl(b.����id, 0) Desc, a.�ⷿid, c.���� "
    Set rsSQL = mdlDrugPacker.OpenSQLRecord(strSQL, "��ȡҩ����ҩ��Ϣ", strNO)
    
    With rsSQL
        If .EOF = False Then
            lngStoreID = NVL(!�ⷿid, 0)
            strReturn = STR_OUT_B & vbCr & _
                        STR_BRID_B & NVL(!����id) & STR_BRID_E & vbCr
            strWin = NVL(!��ҩ����)
        End If
        Do While .EOF = False
            If lngStoreID = NVL(!�ⷿid, 0) Then
                If strWin = "" Then strWin = NVL(!��ҩ����)
                strDrugs = strDrugs & STR_ITEM_B & vbCr & STR_MC_B & NVL(!ҩƷ����) & STR_MC_E & vbCr & STR_ITEM_E & vbCr
            Else
                strReturn = strReturn & _
                            strStore & _
                            STR_YFCK_B & strWin & STR_YFCK_E & vbCr & _
                            STR_YFMX_B & vbCr & strDrugs & STR_YFMX_E & vbCr & _
                            STR_ITEM_E & vbCr
                strDrugs = STR_ITEM_B & vbCr & STR_MC_B & NVL(!ҩƷ����) & STR_MC_E & vbCr & STR_ITEM_E & vbCr
                strWin = NVL(!��ҩ����)
            End If
            
            strStore = STR_ITEM_B & vbCr & STR_YFMC_B & NVL(!ҩ������) & STR_YFMC_E & vbCr
            lngStoreID = NVL(!�ⷿid, 0)
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            strReturn = strReturn & _
                        strStore & _
                        STR_YFCK_B & strWin & STR_YFCK_E & vbCr & _
                        STR_YFMX_B & vbCr & strDrugs & STR_YFMX_E & vbCr & _
                        STR_ITEM_E & vbCr & _
                        STR_OUT_E
            strDrugs = ""
        End If
        
        .Close
        
    End With
    
    PackingWindow_DYEY = strReturn
    Exit Function
    
errHandle:
    strOut = strOut & vbCrLf & "���ء�����ID����ҩҩ������ҩ���ڡ�ҩƷ���ơ���Ϣʧ��"
    PackingWindow_DYEY = ""
End Function

Public Function HaveUpdateFlag() As Boolean
'���ܣ�����Ƿ���Ҫ�ϴ���־����
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2] ANd Rownum < 2"
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "����־", "δ��ҩƷ��¼", "�Ƿ��ϴ�")
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSQL, "����־", "δ��ҩƷ��¼", "�Ƿ��ϴ�")
    End If
    If rsTemp!Rec > 0 Then
        HaveUpdateFlag = True
    End If
    rsTemp.Close
    Exit Function
    
errHandle:
    Call OutputLog("��顰�Ƿ��ϴ����ֶ�ʧ�ܣ�" & Err.Description)
End Function

Public Function UpdateFlag(ByVal lngStockID As Long, ByVal strNO As String) As Boolean
    Dim strSQL As String
    Dim strTmp As String
    Dim arrNO As Variant, arrItem As Variant
    Dim l As Long
    Dim strRecipe As String
    
    On Error GoTo errHandle
    
    strTmp = "�����ϴ���־��ʼ" & vbNewLine
    strTmp = strTmp & "����1��" & lngStockID & "�� ����2��" & strNO & vbNewLine
    
    If Trim(strNO) <> "" Then
        
        If gblnUpdateFlag = False Then
            '�ޡ��Ƿ��ϴ����ֶΣ���������ϴ���־
            UpdateFlag = True
            Exit Function
        Else
            strTmp = strTmp & "��Ҫ���±�־��" & vbNewLine
        End If
    
        arrNO = Split(strNO, "|")
        For l = LBound(arrNO) To UBound(arrNO)
            strRecipe = arrNO(l)
            If strRecipe <> "" Then
                strSQL = "Zl_δ��ҩƷ��¼_�����ϴ���־(" _
                       & lngStockID & "," _
                       & "'" & strRecipe & "')"
                If gintMode = 0 Then
                    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "���±�־")
                Else
                    Call mdlDrugPacker.ExecuteProcedure(strSQL, "���±�־")
                End If
                strTmp = strTmp & arrNO(l) & " ��ɣ�" & vbNewLine
            End If
        Next
    Else
        strTmp = strTmp & "�޵��ݸ��£�"
    End If
    UpdateFlag = True
    Call OutputLog(strTmp & vbCrLf & "�����ϴ���־���")
    Exit Function
    
errHandle:
    strTmp = strTmp & vbNewLine & "�����ϴ���־�쳣:"
    Call OutputLog(strTmp & Err.Description)
End Function
