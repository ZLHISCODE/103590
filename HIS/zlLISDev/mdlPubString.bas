Attribute VB_Name = "mdlPubString"
Option Explicit

'��ģ�鱣�� �ַ��������йصĹ������� ,��ģ���еĺ������� PStr_��ͷ

Public Function PStr_CutCode(ByRef strIn As String, strS As String, strE As String) As String
    '��ָ���Ŀ�ʼ��������������ȡһ���ַ�
    '�ɹ����ؽ�ȡ���ַ���
    Dim lngS As Long, lngE As Long
    lngE = 0: lngS = 0
    lngS = InStr(strIn, strS)
    If lngS > 0 Then lngE = InStr(lngS, strIn, strE)
    PStr_CutCode = ""
    If lngS > 0 And lngE > 0 Then
        PStr_CutCode = Mid(strIn, lngS, lngE - lngS + Len(strE))
        strIn = Mid(strIn, lngE + Len(strE))
    End If
    
End Function

