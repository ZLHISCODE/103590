VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VITEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ΢���� VITEK
'ҽԺ:  ��ú,÷�ӿ�ҽԺ
'����:  ����

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
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '                     a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '                     b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    

    Dim strData As String
    Dim aData() As String
    Dim lngloop As Long
    Dim lng_S As Long, lng_E As Long
    Dim strNow As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H5)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H4)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H3)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H1E)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    If InStr(Replace(strSource, mstrLastInput, ""), Chr(&H1D)) > 0 Then
'        strCmd = Chr(&H6)
'    End If
'    'û���ҵ���ʼλ��ʱ�˳�
'    If Not strSource Like "**mtrsl|iiV2|is*|itSU*" Then
'        strReserved = strSource
'        mstrLastInput = strReserved
'        Exit Sub
'    End If
'
'    'û���ҵ�λ��λ��ʱ�˳�
'    lngStart = InStr(strSource, "mtrsl|iiV2|is")
'    If InStr(lngStart, strSource, "") <= 0 Then
'        strReserved = strSource
'        mstrLastInput = strReserved
'        Exit Sub
'    End If
'    lngEnd = InStr(lngStart, strSource, "") + 2
    '---------------------------------------------------------
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H4)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H1E)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&H1D)) > 0 Then strCmd = Chr(&H6)
    
    mstrCode = mstrCode & Replace(strSource, Chr(&H1E), "")
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        strNow = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        If strNow Like "**mtrsl|iiV2|is*|itSU*" Or strNow Like "**mtrsl|iiV2|is*|itID*" Then
            aData = Split(strNow, "|")
            strData = Format(Mid(aData(7), 3), "yyyy-mm-dd") & "|" & Val(Mid(aData(9), 3)) & "| |ABCDEFGHI|0"
            
            '�������
            For lngloop = 7 To UBound(aData)
                
                If Mid(aData(lngloop), 1, 2) = "o1" Then
                    '�õ�ϸ��
                    strData = Replace(strData, "ABCDEFGHI", Mid(aData(lngloop), 3))
                ElseIf Mid(aData(lngloop), 1, 2) = "a1" Then
                    '�õ����
                    If InStr(",R,S,I,", "," & Trim(UCase(Mid(aData(lngloop + 3), 3))) & ",") > 0 Then
                        strData = strData & "|" & Trim(Mid(aData(lngloop), 3)) & "|" & Trim(UCase(Mid(aData(lngloop + 3), 3))) & "^1^" & Trim(Mid(aData(lngloop + 2), 3))
                    ElseIf InStr(",+,-,", "," & Trim(UCase(Mid(aData(lngloop + 3), 3))) & ",") > 0 Then
                        strData = strData & "|" & Trim(Mid(aData(lngloop), 3)) & "|" & "^^" & Trim(Mid(aData(lngloop + 3), 3))
                    End If
                End If
            Next
            If Trim(strData) <> "" Then
                strResult = strResult & "||" & strData
            End If
            
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub



Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim varItem As Variant, strItem As String, i As Integer, strSampleNO As String
    blnSuccess = True
   
End Function









