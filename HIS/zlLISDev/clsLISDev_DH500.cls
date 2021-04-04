VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DH500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:DH500
'����:�人�д�
'����:�ƽ���
'�û�:�ɶ�����ҽԺ

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
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = "": strValues = ""
    
    If Not (strSource Like "20*/*/*" & Chr(&HD) & "*" & Chr(&HA) & "*") Then
       strReserved = strSource
       Exit Sub
    End If
    
    If InStr(strSource, Chr(&HA)) > 0 Then
        blnEnd = False
        strValues = strSource & Chr(&HA)
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strValues, Chr(&HA)) - InStr(strValues, Chr(&H32))
            strLine = Mid(strValues, InStr(strValues, Chr(&H32)), lngLength)
            strValues = Mid(strValues, InStr(strValues, Chr(&HA)) + 1)       'һ�������Ĵ�������ȡ�꣬��������ݴ浽strReserved���´��ٽ���
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            strLine = Format(Mid(strLine, 1, 16), "yyyy-MM-dd HH:MM") & ":" & Format(Now, "SS") & Mid(strLine, 17)
            strLine = Replace(strLine, "/", "-")
            strLine = Replace(strLine, "(QC-1) ", "| |Ѫ��|1|")
            strLine = Replace(strLine, "(B) ", "| |Ѫ��|0|")
            strLine = Replace(strLine, "(U) ", "| |Ѫ��|0|")
            strLine = Replace(strLine, "(H)", "")
            strLine = Replace(strLine, "(L)", "")
            strLine = Replace(strLine, " ID=", "|")
            strLine = Replace(strLine, "= ", "|")
            strLine = Replace(strLine, "=", "|")
            strLine = Replace(strLine, "(?)", "")
            strLine = Replace(strLine, "< ", "|<")
            strLine = Replace(strLine, "> ", "|>")
            strLine = Replace(strLine, " Na", "|Na")
            strLine = Replace(strLine, " Cl", "|Cl")
            strLine = Replace(strLine, " Ca", "|Ca")
            strLine = Replace(strLine, " pH", "|pH")
            strLine = Replace(strLine, " n-Ca", "|n-Ca")
            strLine = Replace(strLine, " T-Ca", "|T-Ca")
            strResult = Trim(strResult) & "||" & Trim(strLine)
            If InStr(strValues, Chr(&HA)) = 0 Or Len(strValues) < 5 Then
                blnEnd = True
            End If
        Loop
   
        strResult = Replace(strResult, Chr(&HA), "")
        strResult = Replace(strResult, Chr(&HD), "")
        If strResult <> "" Then
           If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        End If
    End If
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function