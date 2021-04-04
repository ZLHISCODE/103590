VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_Roche_ZLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'����:�������������
'�û�:������������ҽԺ
'����:hhd

Option Explicit

Public Function GetStartCmd() As String
'�豸��׼�ӿ�
'���ܣ���ȡ�豸��ʼ��������
'GetStartCmd = ""
    GetStartCmd = Chr(&H6)    ' Chr(6)
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

    Dim iStrPoint As Integer, iStrPointBK As Integer    '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer    '��ǰ������������ַ�λ��

    Dim blnReadItem As Boolean    '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim blnEndSample As Boolean    '�걾����

    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer
    Dim strNowData As String        '��ǰ�ܽ������������ִ�
    Dim aLine() As String
    Dim intLoop As Integer
    Dim str_�걾�� As String
    Dim str_ʱ�� As String
    Dim str_time As String
    Dim str_���� As String
    Const strStartChr As String = ""
    Const strEndChr As String = ""
    
    On Error GoTo errHandle
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False

    'û���ҵ���ʼ���ͽ�����ʱ�˳�
    If InStr(strSource, strStartChr) = 0 Or InStr(strSource, strEndChr) = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If

    strNowData = Mid(strSource, InStr(strSource, strStartChr), InStr(strSource, strEndChr))
    strReserved = Mid(strSource, InStr(strSource, strEndChr) + 1)
    '����
    aLine = Split(strNowData, vbCrLf)

    '�걾��
    str_�걾�� = Val(Trim(Mid(aLine(1), InStr(aLine(1), "Seq.no.") + 7)))

    'ʱ��
    str_ʱ�� = "20" & Mid(aLine(2), 1, 2) & "-" & Mid(aLine(2), 4, 2) & "-" & Mid(aLine(2), 7, 2)
    str_time = Mid(Trim(aLine(2)), 20) & ":00"    'Trim(aLine(2))
    str_���� = str_ʱ�� & " " & str_time
    If Not IsDate(str_����) Then
        str_ʱ�� = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Else
        str_ʱ�� = Format(CDate(str_����), "yyyy-mm-dd hh:mm:ss")
    End If
    strData = str_���� & "|" & str_�걾�� & "| |��Һ|0"

    For intLoop = 3 To UBound(aLine)
        If Trim(aLine(intLoop)) <> "" And Trim(aLine(intLoop)) <> "" And InStr(aLine(intLoop), "Color:") = 0 Then
            strData = strData & "|" & Trim(Replace(Mid(aLine(intLoop), 1, InStr(aLine(intLoop), ":") - 1), "*", "")) & "|" & Trim(Mid(aLine(intLoop), InStr(aLine(intLoop), ":") + 1))
            'strData = strData & "|" & Replace(Replace(Replace(Trim(Mid(aLine(intLoop), 1, 6)), "*", ""), " ", ""), ":", "") & "|" & Trim$(Mid(aLine(intLoop), 7))
        End If
    Next

    strData = Replace(strData, "Brown Yellow", "�ػ�ɫ")
    strData = Replace(strData, "Thin Yellow", "����ɫ")
    strData = Replace(strData, "Yellow", "��ɫ")
    strData = Replace(strData, "Green", "��ɫ")
    strData = Replace(strData, "Primrose", "����ɫ")
    strData = Replace(strData, "Red", "��ɫ")
    strData = Replace(strData, "Orange", "��ɫ")
    strData = Replace(strData, "Blue", "��ɫ")
    strData = Replace(strData, "Brown", "��ɫ")
    strData = Replace(strData, "Unselected", "δѡ")
    strData = Replace(strData, "Other", "����")
    strData = Replace(strData, "Filemot", "�ƺ�ɫ")
    strData = Replace(strData, "Little Dark", "��ģ��")
    strData = Replace(strData, "Dark", "ģ��")
    strData = Replace(strData, "Clear", "�峺")
    strData = Replace(strData, "Feculent", "����")
    '    strData = Replace(strData, "Negative", "����")
    'strData = Replace(strData, "Normal", "����")
    '    strData = Replace(strData, "Positive", "����")
    '    strData = Replace(strData, "Small", "1+")
    '    strData = Replace(strData, "Moderate", "2+")
    '    strData = Replace(strData, "Large", "3+")
    '    strData = Replace(strData, "Trace", "΢��")
    '    strData = Replace(strData, "Thin", "ǳ��ɫ")
    '    strData = Replace$(strData, "Cells/uL", "")
    '    strData = Replace$(strData, "g/L", "")
    '    strData = Replace$(strData, "/ul", "")
    '    strData = Replace$(strData, "mmol/L", "")
    'strData = Replace$(strData, " ", "")

    If Len(strData) > 0 Then
        strResult = strData
    End If
    'strCmd = Chr(&H6&)
    Exit Sub
errHandle:
    WriteLog "Roche_ZLR", "Roche_ZLR", Err.Number & " " & Err.Description
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


