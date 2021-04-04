Attribute VB_Name = "mdlESign"

Option Explicit

Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Public Const G_STR_SPLIT As String = "&&&"

Public Enum CA_TYPE
    CA_����ʡ = 1 '����ʡ����֤����֤����
    CA_����ʡ = 2 '����ʡ����֤����֤����
    CA_������ = 3 '����������֤����֤����
    CA_ɽ��ʡ = 4 'ɽ��ʡ����֤����֤����
    CA_��������ҽԺ = 5 '������Ԫ����֤����֤����,ԭ���ƽ� ��������ҽԺ ����֤����֤����
    CA_����ʡҽԺ = 6 '��Ͷ��������֤����֤����,ԭ���ƽ� ����ʡҽԺ ����֤����֤����
    CA_׼���ҽԺ = 7 '��Ͷ����֤����֤����(����),ԭ���ƽ� ׼���ҽԺ ����֤����֤����,11��12�¸ĳ��ð��ŵ���
    CA_���� = 8
    CA_�㶫 = 9 '�㶫����֤����֤����(����),��û���ò���
    CA_���� = 10 '��������֤����֤����(����)
    CA_����CA�Ĵ� = 11 '��������֤����֤����(�Ĵ�)
    CA_����CA���� = 12 '��������֤����֤����(����),��ʱ���
    CA_����CA���� = 13 '��������֤����֤����(��),��ʱ���
    CA_����CA���� = 14 '��������֤����֤����(����)
    CA_�Ϻ�CA = 15     '�Ϻ�����֤����֤���� ��ʱ���
    CA_����CA = 16     '��������֤����֤����  ��ʱ���
    CA_�½�CA = 17     '�½�����֤����֤����  ��ʱ���(���ں�̩key�ͻ���key������)
    CA_����CA���� = 18 '��������֤����֤����(����_��Ǩ������ҽԺ)
    CA_�ӱ�CA���� = 19 '�ӱ�����֤����֤���� (�����е���ҽԺ)
    CA_����CA���� = 20 '��������֤����֤���� (���������д�Ⱦ��ҽԺ) 20150831
    CA_������Ժ = 21   '����ʡ����֤����֤���� (������Ժ������ҽԺ����ҽ��Ժ)20160321
    CA_���� = 22       '����ʡ����֤����֤�����������޹�˾
    CA_���� = 23       '�����е�������ȫ֤��������޹�˾
    CA_���ְ��� = 24       '����ʡ���ŵ�����֤�������޹�˾
    CA_���ɹ� = 25       '���ɹ�����֤����֤����
    CA_��֤ͨ = 26       '�㶫ʡ����������֤���޹�˾(�����֤ͨ)
End Enum

Public gintCA As CA_TYPE
Public gcnOracle As ADODB.Connection
Public gstrSysName As String
Public glngSys As Long
Public gblnShow As Boolean            '�Ƿ���ʾ���� False-�״���ʾ;True-��ֹ��ʾ
Public gObjFso As New FileSystemObject

Public Type TEST_SIGN
    IsInit      As Boolean      '�Ƿ��ѳ�ʼ��
    strIniFileName As String    '�����ļ���
    
    strSN As String             '���к�
    strUser As String           '�û���
    strPass As String           '����
    strName As String           '����
    dateEnd As String             '֤�鵽��ʱ��
    
    strSignCert As String       '֤������
    strEncCert As String        '����֤������
    strSignImage As String      'ǩ��ͼƬ�ļ���
End Type

Public Type USER_INFO
    strID As String     '�û�ID
    strName As String   '����
    strSignName As String  'ǩ��
    strUserID      As String '���֤��
    lngCertID      As Long  '֤��ID
    strCertSn    As String   '֤�����
    strCert   As String      '֤������
    strCertDN As String      '֤��DN
    strEncCert As String
    strCertID As String
    strPicCode As String    'BASE64ǩ��ͼƬ����
    strSealCode As String   'ǩ��֤��
    strTSCert As String      'ʱ���֤��
    strPicPath As String    'ǩ��ͼƬ����·��
End Type

Public Type PARA_INFO
    strSIGNIP As String      'ǩ��������IP
    strSignPort As String     'ǩ���������˿ں�
    bytSignVersion As Byte    'ǩ���汾 RSA\SM2
    strTSIP As String        'ʱ���������IP
    strTSPort As String       'ʱ����������˿ں�
    strSignURL As String          'ǩ�������ַ
    strTSVersion As String       'ʱ����汾
    blnISTS  As Boolean       '�Ƿ�����ʱ���
    blnIsSign As Boolean      '�Ƿ�����ǩ��������
    blnSignPic As Boolean     '�Ƿ�����ǩ��
    intKeyType As Integer     'Key���� ��������ͬһCA��ͬKEY����� ��:����CA
    strOption  As String      '��ѡ����;�����&��Ϊ�ָ���
End Type
Public Const TEST_MODE = 0      '�Ƿ����ģʽ���� 0- ��1-��

Public mstrCurrPass As String     '��ǰ���� ����ɽ������¼�û���������룬�����û��ظ�����|���ڲ���׮ģ��
Public mstrCurrUser As String     '��ǰ�û� ����ɽ������¼��ǰ�û��������û��ظ�����
Public mUserInfo As USER_INFO      '���浱ǰ����Ա��Ϣ ǩ��ʱ��ʼ��
Public gudtPara As PARA_INFO
Public gstrLogins As String           '����Ѿ�ͨ����¼��֤��key�����к�
Public gobjComLib As Object           '�����������󣬳�ʼ��ʱ��̬����
Public gstrPara  As String
Public glngSign  As Long         '���clsESign��ʵ����Ŀ

Public Const SWP_NOACTIVATE = &H10 '�������
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST As Long = -1
Public Const CON_TOP As Long = 262144
'׼������ʹ����ʼ������ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetSubject(ByVal strSubject As String, ByVal strItem As String) As String
'���ܣ���֤�������л�ȡָ����Ŀ������
    Dim arrItem As Variant
    Dim i As Long
    
    If strSubject <> "" Then
        arrItem = Split(strSubject, ",")
        For i = 0 To UBound(arrItem)
            If UCase(arrItem(i)) Like UCase(strItem) & "=*" Then
                GetSubject = Mid(arrItem(i), InStr(arrItem(i), "=") + 1)
                Exit Function
            End If
        Next
    End If
End Function

Public Function GetSDError(ByVal lngError As Long) As String
'���ܣ�����ɽ��CA�ӿ���������뷵�ش�������
    Dim strError As String
    
    Select Case lngError
        Case 50 'Initcontrol�ӿ�
            strError = "ϵͳ·���Ѿ���ʼ��" '�ô���ɺ���
        Case -1001 'ReadCert�ӿ�
            strError = "֤��·��������ʽ����"
        Case -4100
            strError = "֤��û�п�ͨӦ�û���Ӧ���ļ����𻵻����̡�EKEYû�в��"
        Case -4001
            strError = "֤��δ��Ӧ���ڻ�ϵͳʱ�����ô���"
        Case -4002
            strError = "֤��û�п�ͨ���Ӧ��"
        Case -4003
            strError = "֤��İ�ȫӦ���ѹ��ڻ�ϵͳʱ�����ô���"
        Case -4004
            strError = "֤�����������ļ���ƥ��"
        Case 51
            strError = "֤���ѹ��ڻ�ϵͳʱ�����ô���"
        Case -9018
            strError = "û���ҵ�֤��"
        Case -9005
            strError = "֤����������֤�鲻����"
        Case -1003
            strError = "EKEYδ���"
        Case -9003
            strError = "���ܷ��ʼ����豸" '�������
        Case -5101
            strError = "����ʵ����EKEY֤��"
        Case -2001 '�����ŷ⼰��ȡͼ�½ӿ�
            strError = "��֤�������"
        Case -3000
            strError = "����ͼ���ļ��ײ��ʧ��"
        Case -3001
            strError = "��ȡͼ���ļ�ʧ��"
        Case -3002
            strError = "����ͼ���ļ�ʧ��"
        Case -3004
            strError = "ҪǶ����ַ����ȳ���ͼƬ�����ɳ���"
        Case -3006
            strError = "��ѹʧ��"
        Case -3009
            strError = "���ļ�����"
        Case -5001
            strError = "����Base64����"
        Case -5002
            strError = "����Base64����"
        Case -9009 '�������������б�
            strError = "CRL������"
        Case -9012
            strError = "֤����������"
        Case -9014
            strError = "��֤����Ч"
        Case -9021
            strError = "˽Կ������"
        Case -9022
            strError = "�㷨����Կ��ƥ��"
        Case -9026
            strError = "֤����㷨��ƥ��"
        Case -9027
            strError = "ǩ��ʧ��"
        Case -9028
            strError = "��֤ǩ��ʧ��"
        Case -9029
            strError = "����ʧ��"
        Case -9030
            strError = "����ʧ��"
        Case -9043
            strError = "�����ļ�������"
        Case Else
            strError = "δ֪����"
    End Select
    
    GetSDError = strError
End Function

Public Function PackBytes(ByVal strData As String) As Byte()
'���ܣ����ַ���ת��Ϊ�ֽ�����
    Dim arrByte() As Byte
    Dim intAscii As Integer, intIdx As Integer
    Dim strChar As String, strHex As String
    Dim i As Integer
    
    If strData = "" Then Exit Function
    ReDim arrByte(LenB(strData) - 1)
    
    intIdx = 0
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        If strChar = Space(1) Then strChar = "+" '�ո�ת��
        
        If strChar <> "" Then
            intAscii = Asc(strChar)
            If intAscii >= 0 Then
                arrByte(intIdx) = Asc(strChar)
                intIdx = intIdx + 1
            Else
                'Ascii<0��Ϊ����,�ָߵ��ֽ�ת��Byte������
                strHex = Hex(intAscii)
                arrByte(intIdx) = Val("&H" & left(strHex, 2))
                arrByte(intIdx + 1) = Val("&H" & Right(strHex, 2))
                intIdx = intIdx + 2
            End If
        End If
    Next
    ReDim Preserve arrByte(intIdx - 1) '�ص�����Ĳ���
    
    PackBytes = arrByte
End Function

Public Function IsUpdateRegCert(udtUser As USER_INFO, ByVal strDate As String, ByRef blnReDo As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���ܣ�ǩ����ȡ��ǩ��ʱ���øú�����֤��󶨣�
'������udtUser-��ǰȡֵ���°��û���Ϣ
'      strDate-��֤�����Ч����
'      blnRedo-True:���ע���,False-δ���ע���
'����:��ΰ��
'����:2015-09-10
'--------------------------------------------------------------------------------------------------------------------
    Dim strTmp  As String
    Dim intDay As Integer
    Dim blnDo As Boolean
    Dim blnTip As Boolean
    Dim strMsg As String
    Dim strFileName As String
    Dim blnTrans As Boolean
    
    If mUserInfo.strCertSn <> udtUser.strCertSn Then
        '���֤��һ������������ǩ��ֵ��� �������,����ע��
        If mUserInfo.strUserID = udtUser.strUserID And (mUserInfo.strName = udtUser.strName Or mUserInfo.strSignName = udtUser.strName) Then
            '���ɵ���Ч��
            If strDate <> "" Then
            '��֤�ͻ���֤����Ч��ʣ������
                If Not gintCA = CA_�Ϻ�CA Then
                    intDay = CheckValidaty(CDate(strDate))
                Else
                    intDay = Val(strDate)
                End If
                
                If (intDay > 0) Then
                    'δ����
                    strTmp = vbCrLf & "����֤�黹��" & intDay & "�����,��������һ���µ�֤���Ƿ�����ע�᣿" & vbCrLf
                    strTmp = strTmp & "ע��:���ע�����µ�,�Ժ�ֻ��ʹ���µ�֤�飡"
                    If MsgBoxEx(strTmp, vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnTip = True
                    Else
                        blnDo = False
                    End If
                ElseIf (intDay <= 0) Then
                    '����
                    strTmp = "����֤���ѹ��� " & Abs(intDay) & " ��,��������һ���µ�֤���Ƿ�����ע�᣿"
                    If MsgBoxEx(strTmp, vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnTip = True
                    Else
                        blnDo = False
                    End If
                End If
               
                If blnTip Then
                    strTmp = "�ҵ���Ϣ��" & vbCrLf
                    strTmp = strTmp & Space(2) & "����:" & IIf(mUserInfo.strSignName = "", mUserInfo.strName, mUserInfo.strSignName) & vbTab & "���֤�ţ�" & mUserInfo.strUserID
                    strTmp = strTmp & vbCrLf
                    strTmp = strTmp & "֤����Ϣ��" & vbCrLf
                    strTmp = strTmp & Space(2) & "ʹ���ߣ�" & udtUser.strName & vbTab & "�����֤��:" & udtUser.strUserID & vbCrLf
                    strTmp = strTmp & "�˶���Ϣ����,��""ȷ��""���ע�ᡣ" & vbCrLf
                    strTmp = strTmp & "�˶���Ϣ����,��""ȡ��""�ݲ�ע�ᡣ" & vbCrLf
                    If MsgBoxEx(strTmp, vbOKCancel + vbInformation + vbDefaultButton1, "ע����Ϣ") = vbOK Then
                        blnDo = True
                    Else
                        blnDo = False
                    End If
                End If
            Else
                blnDo = False   'ϵͳע��KEY����Ч����ʱ��Ϊ��ʱ,�ݲ�����
            End If
            
            If blnDo Then
                'ע��
                strTmp = "zl_��Ա֤���¼_Insert(" & mUserInfo.strID & "," & _
                            "'" & Replace(udtUser.strCertDN, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strCertSn, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strCert, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strEncCert, "'", "''") & "'," & _
                            "'" & Replace(udtUser.strTSCert, "'", "''") & "')"
                On Error GoTo errH
                
                '����ǩ��ͼƬ
                If gintCA = CA_���ְ��� Then
                    udtUser.strPicPath = ANXIN_GetSeal   'ǩ��ͼƬ��ʱ,��ֻ�ڸ���ǩ��ͼƬ��ʱ����ȡ
                End If
                If udtUser.strPicCode <> "" Then
                    strFileName = SaveBase64ToFile("gif", udtUser.strUserID, udtUser.strPicCode)
                Else
                    strFileName = udtUser.strPicPath
                End If
                gcnOracle.BeginTrans: blnTrans = True
                If strFileName <> "" Then
                    If SaveSignPIC(Val(mUserInfo.strID), strFileName) = False Then
                        GoTo errH
                    End If
                End If
                Call gobjComLib.zlDatabase.ExecuteProcedure(strTmp, gstrSysName)
                If udtUser.strSealCode <> "" Then
                    If Not gobjComLib.Sys.Savelob(100, 14, mUserInfo.strID & "," & udtUser.strCertSn, udtUser.strSealCode, 1) Then
                        blnTrans = True
                        GoTo errH
                    End If
                End If
                gcnOracle.CommitTrans: blnTrans = False
                strMsg = "֤����³ɹ���"
                blnDo = False
                blnReDo = True
            Else
                blnDo = False
            End If
        Else
            strTmp = "ע����Ϣ��" & vbCrLf
            strTmp = strTmp & Space(2) & "����:" & IIf(mUserInfo.strSignName = "", mUserInfo.strName, mUserInfo.strSignName) & vbTab & "���֤�ţ�" & mUserInfo.strUserID
            strTmp = strTmp & vbCrLf
            strTmp = strTmp & "֤����Ϣ��" & vbCrLf
            strTmp = strTmp & Space(2) & "ʹ���ߣ�" & udtUser.strName & vbTab & "�����֤��:" & udtUser.strUserID & vbCrLf
            strTmp = strTmp & "��ǰ֤����ע����Ϣ��һ��,����ʹ��!" & vbCrLf
            strMsg = strTmp
            blnDo = False
        End If
    Else
        blnDo = True '������һ��������ǩ��/ȡ��ǩ����
    End If
    
    If strMsg <> "" Then
        MsgBoxEx strMsg, vbOKOnly + vbInformation, gstrSysName
    End If
    IsUpdateRegCert = blnDo
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    IsUpdateRegCert = False
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Private Function SaveSignPIC(ByVal lng��Աid As Long, ByVal strFileName As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle
    blnOk = gobjComLib.Sys.Savelob(100, 15, lng��Աid, strFileName)
    SaveSignPIC = blnOk
    Exit Function
ErrHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function MsgBoxEx(prompt, Optional ByVal buttons As Long, Optional ByVal title As String) As Long
'����:��262144 ʼ����Msgboxλ�������ʾ���ƶ��������ʱ��ʾ��δ�ö���ʾ��
    MsgBoxEx = MsgBox(prompt, CON_TOP Or buttons, title)
End Function

