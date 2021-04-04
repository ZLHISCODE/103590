VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_STACompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:STACompact��Ѫ���ǣ�STAGO
'����:���� STAGO
'����:�¶�
'�û�:������ɽҽԺ,����ʡ����ҽԺ

Private mstrCode As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim lngS As Long, lngE As Long, strCode As String, varTmp As Variant, iRow As Integer, strLine As String
    Dim strDate As String, strSampleNo As String, strItem As String, strValue As String, strItems As String
    On Error GoTo hErr
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    mstrCode = mstrCode & strSource
    lngE = 0: lngS = InStr(mstrCode, Chr(&H5))
    If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid(mstrCode, lngE + 1)
        varTmp = Split(strCode, Chr(&HD) & Chr(&HA) & Chr(&H2))
        
        strDate = "": strSampleNo = "": strItems = ""
        
        For iRow = LBound(varTmp) To UBound(varTmp)
            strLine = Trim$("" & varTmp(iRow))
            If Left(strLine, 5) = Chr(&H5) & Chr(&H2) & "1H|" Then
                strDate = Split(strLine, "|")(13)
                If Len(strDate) >= 14 Then
                    strDate = Mid$(strDate, 1, 4) & "-" & Mid$(strDate, 5, 2) & "-" & Mid$(strDate, 7, 2) & " " & Mid$(strDate, 9, 2) & ":" & Mid$(strDate, 11, 2) & ":" & Mid$(strDate, 13, 2)
                End If
                If Not (strDate <> "" And IsDate(strDate)) Then strDate = ""
            ElseIf strLine Like "?O|1|*|R*" Then
                strSampleNo = Trim$("" & Split(strLine, "|")(2))
            ElseIf strLine Like "?O|1|*|S*" Then
                strSampleNo = Trim$("" & Split(strLine, "|")(2))
                If strSampleNo <> "" Then strSampleNo = strSampleNo & "^1^"
            ElseIf strLine Like "?R|*|^^^*" Then
                strItem = Trim(Replace$(Split(strLine, "|")(2), "^^^", ""))
                strValue = Trim(Split(strLine, "|")(3))
                If strItem <> "" And strValue <> "" Then
                    strItems = strItems & "|" & strItem & "|" & strValue
                End If
            End If
        Next
        
        If strDate <> "" And strSampleNo <> "" And strItems <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Ѫ��|0" & strItems
        End If
        
        lngE = 0: lngS = InStr(mstrCode, Chr(&H5))
        If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "STACompact", "��" & CStr(Erl()) & "��", Err.Description
    'Resume
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function


