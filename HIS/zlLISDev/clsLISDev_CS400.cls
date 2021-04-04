VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  �³�XC8001����������
'ҽԺ:  ��������ҽԺ
'����:  ������
'ʱ��:  2016/01/10

Private mstrCode As String '������յ�������

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
    Dim intloop As Integer
    Dim strDate As String               '����
    Dim strSampleNO As String           '�걾��
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    Dim str_Sample As String
    Dim str_No As String, Str_Item As String, str_Value As String
    Dim varLine As Variant, varData As Variant, iRow As Integer, varData1 As Variant
    
    On Error GoTo errhand
    
    '��һ����������յ�������
    '��������Ϊ����ͨѶ�ٶȽ�������������ʱ���п��ܻ���ն�β��ܽ��յ�һ��
    '����������
    If InStr(strSource, Chr(&H5)) > 0 Then
    strCmd = Chr(6)
     ElseIf InStr(strSource, vbCrLf) > 0 Then
    strCmd = Chr(6)
     ElseIf InStr(strSource, Chr(&H4)) > 0 Then
    strCmd = Chr(6)
    End If
    mstrCode = mstrCode & strSource
   
    '�ڶ������ж��Ƿ��յ���һ�����������ݣ������������ݲſ�ʼ���룬
    '����ʲô��������
'    mstrCode = Replace(mstrCode, Chr(13), "")
'    mstrCode = Replace(mstrCode, Chr(10), "")
'    mstrCode = Replace(mstrCode, Chr(28), "")
    lngBegin = InStr(mstrCode, "H|\^&|")
    lngEnd = InStr(lngBegin, mstrCode, "L|1|N")
   
    Do While lngBegin > 0 And lngEnd > 0
        '���������յ����������ݣ��Ȱ�������ݴӻ�����ȡ������
        str_Sample = Mid(mstrCode, lngBegin, lngEnd - lngBegin + 10)
        mstrCode = Mid(mstrCode, lngEnd + 12)
 
        '���Ĳ���ȡ�����ڣ��걾�ţ���Ŀ��������
        varLine = Split(str_Sample, Chr(10))
        
        For iRow = 0 To UBound(varLine)
            If varLine(iRow) Like "H|\^&|*" Then
                varData = Split(varLine(iRow), "|")
                If varData(12) <> "" Then
                    strDate = Mid(varData(12), 1, 4) & "-" & Mid(varData(12), 5, 2) & "-" & Mid(varData(12), 7, 2) & " " & Mid(varData(12), 9, 2) & ":" & Mid(varData(12), 11, 2)
                End If
            End If
            
'            If strSampleNO = "" And varLine(iRow) Like "?P|*|*" Then
'                varData = Split(varLine(iRow), "|")
'                strSampleNO = Val(varData(1))
'            End If
            
            If varLine(iRow) Like "?O|*|*" Then
                varData = Split(varLine(iRow), "|")
                strSampleNO = Val(Split(varData(2), "^")(1))
            End If
            
            If varLine(iRow) Like "*R*^^^*" Then
                varData = Split(varLine(iRow), "|")
                If varData(3) <> "No Result" Then
                    varData1 = Split(varData(2), "^")
                    Str_Item = Str_Item & "|" & varData1(3) & "|" & varData(3)
                    If strDate = "" Then
                        strDate = Mid(varData(12), 1, 4) & "-" & Mid(varData(12), 5, 2) & "-" & Mid(varData(12), 7, 2) & " " & Mid(varData(12), 9, 2) & ":" & Mid(varData(12), 11, 2)
                    End If
                End If
            End If
        Next
           
        '���岽 ��ȡ����������֯�ɹ涨��ʽ����strResult������
        If Str_Item <> "" And Val(strSampleNO) <> 0 Then
            strResult = strResult & "||" & strDate & "|" & strSampleNO & "| |Ѫ��|0" & Str_Item
                        
        End If
        Str_Item = ""
        strSampleNO = ""
        strDate = ""
        lngBegin = 0
        lngEnd = 0
        '������ ��黺���е����ݣ�������������ļ�¼�����������
        lngBegin = InStr(mstrCode, "H|\^&|")
        If lngBegin > 0 Then lngEnd = InStr(lngBegin, mstrCode, "L|1|N")
    Loop
    '������ɣ��������||�ַ�ȥ����
    If strResult <> "" Then strResult = Mid(strResult, 3)
    strReserved = mstrCode
    Exit Sub
errhand:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
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
    
End Function



