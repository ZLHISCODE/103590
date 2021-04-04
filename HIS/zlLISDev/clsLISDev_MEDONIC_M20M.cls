VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MEDONIC_M20M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:MEDONIC M20 Ѫϸ��������
'����:��������Ƽ����޹�˾
'����:���
'�û�:�»�
'����:2010-01-29

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_ʱ�� As String, str_������ As String, str_�걾 As String
    Dim str_��Ŀ As String, str_��� As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String, str_Img As String, strImgData As String
    Dim varline As Variant, blnEnd As Boolean
    Dim i As Integer, str_Line As String
    Dim str_smpInfo As String, str_smpresults As String
    Dim str_hgrams As String, str_hgram As String
    Dim lngLength As Long
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*<sample>*</sample>*" Then
        blnEnd = False
        strReserved = strSource
        
        Do While Not blnEnd
            lngLength = InStr(strReserved, "</sample>") + 9 - InStr(strReserved, "<sample>")
            str_Data = Mid(strReserved, InStr(strReserved, "<sample>"), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, "</sample>") + Len("</sample>"))
            
            str_��Ŀ = "": str_������ = "": str_ʱ�� = ""
            str_Img = "": strImgData = ""
            '<smpinfo></smpinfo>
            str_smpInfo = GetKeyVal(str_Data, "smpinfo")
            varline = Split(str_smpInfo, Chr(&HD) & Chr(&HA))
            
            For i = LBound(varline) To UBound(varline)
                str_Tmp = GetKeyVal(varline(i), "n")
                If str_Tmp = "SEQ" Then
                    str_������ = GetKeyVal(varline(i), "v")
                ElseIf str_Tmp = "DATE" Then
                    str_ʱ�� = Replace(GetKeyVal(varline(i), "v"), "T", " ")
                End If
                If str_������ <> "" And str_ʱ�� <> "" Then Exit For
            Next
            
            '<smpresults></smpresults>
            str_smpresults = GetKeyVal(str_Data, "smpresults")
            varline = Split(str_smpresults, Chr(&HD) & Chr(&HA))
            For i = LBound(varline) To UBound(varline) - 1
                If varline(i) <> "" And GetKeyVal(varline(i), "n") <> "" And GetKeyVal(varline(i), "v") <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & GetKeyVal(varline(i), "n") & "|" & GetKeyVal(varline(i), "v")
                End If
            Next
            str_��Ŀ = Replace(str_��Ŀ, Chr(&HD) & Chr(&HA), "")
            
            '<hgrams></hgrams>
            str_hgrams = GetKeyVal(str_Data, "hgrams")
            For i = 0 To 2
                str_hgram = GetKeyVal(str_Data, "hgram")
                str_Data = DelStrRead(str_Data, "hgram")
                str_Tmp = GetKeyVal(str_hgram, "n")
                    
                If str_Tmp = "PLT" Then
                    str_Img = str_Img & "^PLT;0;0,80;"
                ElseIf str_Tmp = "RBC" Then
                    str_Img = str_Img & "^RBC;0;0,80;"
                ElseIf str_Tmp = "WBC" Then
                    str_Img = str_Img & "^WBC;6;0,80;"
                End If
        

                Do While GetKeyVal(str_hgram, "v") <> ""
                    strImgData = GetKeyVal(str_hgram, "v")
                    str_hgram = DelStrRead(str_hgram, "v")
                    'ȥ���ո�
                    strImgData = Replace(strImgData, " ", ";")
                    'ȡ���س�����
                    strImgData = Replace(strImgData, Chr(&HD) & Chr(&HA), ";")
                    strImgData = Mid(strImgData, 1, InStrRev(strImgData, ";") - 1)
                    strImgData = Mid(strImgData, InStr(strImgData, ";") + 1)
                    If str_Tmp = "WBC" Then strImgData = "~" & strImgData
                    str_Img = str_Img & strImgData
                Loop
            Next
            
            If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |Ѫ��|0" & str_��Ŀ
                If str_Img <> "" Then strResult = strResult & vbCrLf & Mid(str_Img, 2)
            End If
            
            If Not strReserved Like "*<sample>*</sample>*" Then
                blnEnd = True
            End If
        Loop
    Else
       strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
End Sub

Private Function GetKeyVal(ByVal strTemp As String, ByVal strKey As String) As String
    '���� key ��ȡ VAL
    Dim int_Start As Integer
    Dim int_Length As Integer
    Dim str_KeyVal As String
    
    int_Start = InStr(strTemp, "<" & strKey & ">") + IIf(InStr(strTemp, "<" & strKey & ">") > 0, Len("<" & strKey & ">"), 0)
    If int_Start = 0 Then Exit Function
    int_Length = InStr(strTemp, "</" & strKey & ">") - int_Start
    str_KeyVal = Mid(strTemp, int_Start, int_Length)
    GetKeyVal = str_KeyVal
End Function

Private Function DelStrRead(ByVal strTemp As String, ByVal strKey As String) As String
    'ɾ���Ѷ�ȡ�����ַ���
    Dim int_Start As Integer
    Dim int_Length As Integer
    Dim str_Result As String
    
    int_Start = InStr(strTemp, "</" & strKey & ">") + Len("</" & strKey & ">")
    If int_Start = 0 Then Exit Function
    str_Result = Mid(strTemp, int_Start)
    DelStrRead = str_Result
End Function

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