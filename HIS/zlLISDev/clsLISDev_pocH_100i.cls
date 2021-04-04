VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_pocH_100i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  pocH-100i ������Ѫ���� �ձ�����sysmex
'ҽԺ:  ������������ҽԺ
'����:  �¶�
'ʱ��:  2008/09/05

Private mstrCode As String
Private mLastReturn As String

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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim str_Code As String
    Dim str_Line As String               '��ǰҪ�������ִ�
    Dim lng_S As Long                  '��ʼλ��
    Dim lng_E As String                '����λ��
    
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String
    Dim StrWbc As String, StrRbc As String, StrPLt As String
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    
    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = 0
    End If

    
        
    Do While lng_S > 0 And lng_E > 0

        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_���� = "": str_�걾�� = "": str_��Ŀ = ""
        If UCase(Mid(str_Code, 2, 3)) = "D1U" Then
            str_���� = Trim(Mid(str_Code, 45, 4) & "-" & Mid(str_Code, 49, 2) & "-" & Mid(str_Code, 51, 2))
            
            str_�걾�� = Val(Mid(str_Code, 54, 15))
            
            str_��Ŀ = "|WBC|" & GetValues(Mid(str_Code, 76, 4), 1)
            str_��Ŀ = str_��Ŀ & "|RBC|" & GetValues(Mid(str_Code, 81, 4), 2)
            str_��Ŀ = str_��Ŀ & "|HGB|" & GetValues(Mid(str_Code, 86, 4), 0)
            str_��Ŀ = str_��Ŀ & "|HCT|" & GetValues(Mid(str_Code, 91, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MCV|" & GetValues(Mid(str_Code, 96, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MCH|" & GetValues(Mid(str_Code, 101, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MCHC|" & GetValues(Mid(str_Code, 106, 4), 0)
            str_��Ŀ = str_��Ŀ & "|PLT|" & GetValues(Mid(str_Code, 111, 4), 0)
            str_��Ŀ = str_��Ŀ & "|LYM%|" & GetValues(Mid(str_Code, 116, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MXD%|" & GetValues(Mid(str_Code, 121, 4), 1)
            str_��Ŀ = str_��Ŀ & "|NEUT%|" & GetValues(Mid(str_Code, 126, 4), 1)
            str_��Ŀ = str_��Ŀ & "|LYM#|" & GetValues(Mid(str_Code, 131, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MXD#|" & GetValues(Mid(str_Code, 136, 4), 1)
            str_��Ŀ = str_��Ŀ & "|NEUT#|" & GetValues(Mid(str_Code, 141, 4), 1)
            str_��Ŀ = str_��Ŀ & "|RDW-SD|" & GetValues(Mid(str_Code, 146, 4), 1)
            str_��Ŀ = str_��Ŀ & "|RDW-CV|" & GetValues(Mid(str_Code, 151, 4), 1)
            str_��Ŀ = str_��Ŀ & "|PDW|" & GetValues(Mid(str_Code, 156, 4), 1)
            str_��Ŀ = str_��Ŀ & "|MPV|" & GetValues(Mid(str_Code, 161, 4), 1)
            str_��Ŀ = str_��Ŀ & "|P-LCR|" & GetValues(Mid(str_Code, 166, 4), 1)
            
            If Val(str_�걾��) <> 0 And IsDate(str_����) And str_��Ŀ <> "" Then
                mLastReturn = str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0"
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0" & str_��Ŀ
            End If
            
        ElseIf UCase(Mid(str_Code, 2, 2)) = "D2" Then
            If mLastReturn <> "" Then
                str_Code = Mid(str_Code, 4)
                StrWbc = Left(str_Code, 100)
                StrRbc = Mid(str_Code, 101, 100)
                StrWbc = HexToData(StrWbc)
                StrRbc = HexToData(StrRbc)
                
                If mLastReturn <> "" Then
                    strResult = strResult & "||" & mLastReturn & vbCrLf & "WBC;0;280,40,0-0,9-,18-100,27-,36-200" & StrWbc
                    strResult = strResult & "^" & "RBC;0;280,50,0-0,10-50,20-100,30-150,40-200" & StrRbc
                End If
            End If
        ElseIf UCase(Mid(str_Code, 2, 2)) = "D3" Then
            If mLastReturn <> "" Then
                StrPLt = Mid(str_Code, 4, 70)
                StrPLt = HexToData(StrPLt)
                
                strResult = strResult & "||" & mLastReturn & vbCrLf & "PLT;0;280,38,0-0,10-10,20-20,30-30" & StrPLt
                
                mLastReturn = ""
            End If
        End If
        
        
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = 0
        End If
        
    Loop

    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
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

Private Function HexToData(ByVal strIn As String) As String
    Dim i As Integer, strData As String
    
    For i = 1 To Len(strIn) Step 2
        strData = strData & ";" & CByte("&H" & Mid(strIn, i, 2)) * 2
    Next
    HexToData = strData
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '����:��ʽ������ֵ
    'str_in:����ʽ�����ַ���
    'int_i:�����������ʱ,������С��λ��
    
    If str_in = "*000" Then
        GetValues = "----"
    ElseIf str_in = "     " Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function
