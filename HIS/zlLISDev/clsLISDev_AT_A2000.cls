VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AT_A2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��ͼ��ѧ�Ź��� A2000
'����:��ͼ
'����:���
'�û�:��ͨ������ҽԺ
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
    Dim str_Code As String
    Dim lng_S As Long, lng_E As Long, str_��Ŀ As String, str_�걾�� As String, str_���� As String
    Dim str_Line As String, i As Integer, str_Tmp As String
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H7B))

    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H7D))
    Else
        lng_E = InStr(mstrCode, Chr(&H7D))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S + 1, lng_E - lng_S - 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------
        str_��Ŀ = "": str_�걾�� = "": str_���� = ""
        
        str_�걾�� = Mid(Trim(Split(str_Code, ",")(2)), 4)
        str_���� = Format("" & Now(), "yyyy-MM-dd HH:mm:ss")

        str_��Ŀ = str_��Ŀ & "|" & Mid(Trim(Split(str_Code, ",")(3)), 4) & "|" & Replace(Trim(Split(str_Code, ",")(6)), "[S]", "")

        If Val(str_�걾��) > 0 And str_��Ŀ <> "" And str_���� <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_��Ŀ
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, Chr(&H7B))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H7D))
        Else
            lng_E = InStr(mstrCode, Chr(&H7D))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
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


