VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_TBA_1200FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:TBA-1200FR ����������
'����:��֥
'����:�¶�
'�޸���:������,˫��ӿ�
'�޸�ʱ�䣺2013-11-11
'�û�:��������ҽԺ,�»�����ҽԺ
Private mstrCode As String

Private mstrRes As String

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
    
    On Error GoTo errH
    If InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H2) & Chr(&H6) & Chr(&H3)
'    ElseIf InStr(strSource, Chr(&H2)) > 0 Then
'        strCmd = Chr(&H2) & Chr(&H6) & Chr(&H3)
    End If
    If strSource <> "" Then mstrCode = mstrCode & strSource
'    Debug.Print strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
   
    Do While lng_S > 0 And lng_E > 0
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------
        If str_Code Like "*" & Chr(&H2) & "R *" Then                                            '�������
            str_��Ŀ = "": str_�걾�� = "": str_���� = ""
            
            str_�걾�� = Trim(Mid(str_Code, 8, 8))
            str_���� = Mid(str_Code, 35, 12)
            
            If IsNumeric(str_����) Then
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                If Not IsDate(str_����) Then str_���� = ""
            Else
                str_���� = ""
            End If
            
            str_Line = str_Code
            
            If InStr(str_Line, Chr(&H17)) > 0 Then str_Line = Mid(str_Line, 1, InStr(str_Line, Chr(&H17)))
            str_Line = Mid(str_Line, 50)
            
            For i = 0 To Len(str_Line) / 15
                If Len(str_Line) < 15 Then Exit For
                str_Tmp = Trim(Mid(str_Line, 5, 6))
                If str_Tmp <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 4)) & "|" & str_Tmp
                End If
                str_Line = Mid(str_Line, 16)
            Next
            '�����ʿ�
            If str_��Ŀ <> "" And Val(str_�걾��) <= 0 Then
                str_�걾�� = 10001
            End If
            
            If Val(str_�걾��) > 0 And str_��Ŀ <> "" And str_���� <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_��Ŀ
            End If
        ElseIf str_Code Like "*" & Chr(&H2) & "Q *" Then                                        '˫��ͨѶ
            str_��Ŀ = "": str_�걾�� = "": str_���� = ""
            str_�걾�� = Trim(Mid(str_Code, 8, 20))
            str_���� = Format(Now, "YYYY-MM-DD hh:mm:ss")
            If Len(str_�걾��) = 8 Then
                If Val(str_�걾��) <> 0 Then
                    strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��)
                End If
            ElseIf Len(str_�걾��) = 14 Then
                If Val(str_�걾��) <> 0 Then
                    strResult = strResult & "||" & str_���� & "|^^" & Val(str_�걾��)
                End If
            End If

            If strResult <> "" Then
                strCmd = "1|"                                                         '˫��ָ��
                mstrRes = str_Code
            End If
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop

    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    Call WriteLog("TBA_1200FR.SendSample ����: ", Err.Number, Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ���������
    '   ��3��Ԫ�أ��걾����
    '   ��4��Ԫ�أ������־
    '   ��5��Ԫ�أ���������
    '   ��6��Ԫ�أ��̺ţ�����
    '   ��7��Ԫ�أ�����ID^�Ա�^��������^����^����ȫƴ^ϡ�ͱ���
    '   ��8��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNo As String, iPos As Long
    
    On Error GoTo errHandle
    
    blnSuccess = True
    If strResponse <> "" Then
        strResponse = mstrRes
    End If
    If strResponse Like "*Q*" Or Len(Trim(strResponse)) = 0 Then    '��������
        If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
            '�������
            '�����豸��
            aItem = Split(strSample, "|")
            '�걾��SampleID
            strBuffer = aItem(1)
            strSampleID = aItem(5)
            
            '��ȡ�����ļܺźͱ���
            strRackNo = ""
            
            
            If Len(Trim(strResponse)) > 0 Then
                strRackNo = Mid(strResponse, 28, 7)
            End If
            
            If Len(Trim(strSampleID)) > 0 Then
                'SampleID
                SendSample = "O " & Right("00000000000000" & strSampleID, 14) & Space(6) & strRackNo & "  1"
            Else
                SendSample = "O " & Right("00000000" & strBuffer, 8) & Space(12) & strRackNo & "  1"
            End If
            
            For i = 10 To UBound(aItem)
                SendSample = SendSample & Right(Space(4) & aItem(i), 4) & "1"
            Next
            
            SendSample = SendSample & ""
            Call WriteLog("TBA_1200FR", "TBA_1200FR.SendSample ���͵���������: ", SendSample)
            iSendStep = 0
        Else
            SendSample = Chr(&H2) & Chr(&H6) & Chr(&H3)
            iSendStep = iSendStep + 1
        End If
    Else
        SendSample = Chr(&H2) & Chr(&H6) & Chr(&H3)
        iSendStep = 0
        blnSuccess = False    '����ʧ��
    End If
        
    Exit Function

errHandle:
    Call WriteLog("TBA_1200FR.SendSample ����: ", Err.Number, Err.Description)
End Function




