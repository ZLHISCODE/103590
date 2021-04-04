VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_2020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���� 2020
'����:����
'����:����
'�û�:����

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, strimg As String
    Dim lngS As Long, lngE As Long
    Dim strNow As String
    Dim aLine() As String, aItem() As String
    Dim intLoop As Integer
    Dim strData As String
    
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
'    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    mstrCode = mstrCode & strSource
    
    'û�п�ʼ��ʱ�˳�
    If InStr(mstrCode, Chr(&H5)) = 0 Then
        Exit Sub
    End If
    
    'û�н�����ʱ�˳�
    If InStr(mstrCode, Chr(&H4)) = 0 Then
        Exit Sub
    End If
    
    strNow = Mid(mstrCode, InStr(mstrCode, Chr(&H5)))
    strNow = Mid(mstrCode, 1, InStr(mstrCode, Chr(&H4)))
    
    mstrCode = Mid(mstrCode, InStr(mstrCode, Chr(&H4)) + 1)
    
    '��ʼ����
    aLine = Split(strNow, vbCrLf)
    
    For intLoop = 0 To UBound(aLine)
        If InStr(aLine(intLoop), "|") > 0 Then
            aItem = Split(aLine(intLoop), "|")
            strTmp = aItem(0)
            strTmp = Replace(Replace(strTmp, vbCr, ""), vbLf, "")
            If Right(strTmp, 1) = "O" Then
                '�걾�ź�ʱ��
                strData = Mid(aItem(6), 1, 4) & "-" & Mid(aItem(6), 5, 2) & "-" & Mid(aItem(6), 7, 2) & " " & _
                                Mid(aItem(6), 9, 2) & ":" & Mid(aItem(6), 11, 2) & ":" & Mid(aItem(6), 13, 2) & _
                          "|" & aItem(2) & "| |ѪҺ|0"
            End If
            
            If Right(strTmp, 1) = "R" Then
                '���
                strData = strData & "|" & Split(aItem(2), "^")(3) & "|" & aItem(3)
            End If
        End If
    Next
    strData = Replace(strData, "||", "| |")
    strResult = strData
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
    
End Function



