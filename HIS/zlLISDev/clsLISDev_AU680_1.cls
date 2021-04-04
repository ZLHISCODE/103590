VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU680_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'����:AU680
'����:
'����:������
'�û�:��������ҽԺ
Private mstrCode As String, mstrWaitSend As String, strRackNo As String

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
    
    Dim str_No As String, str_Date As String '�걾���Լ�ʱ��
    Dim str_Item As String, str_Value As String '��Ŀ�Լ����
    Dim m As Long, n As Long, lngA As Long '�������ʱ��ѭ����
    Dim str_Hc As String '�洢��ʱ���ݣ�ʹ������������
    Dim str_Jz As String '�����־
    Dim str_���� As String '����
    Dim str_Tm As String
    
    str_Jz = ""
    
    '��������������������
    mstrCode = mstrCode & strSource
'    If mstrCode Like "*" & Chr(&H20) & Chr(&H3) Then
'        strCmd = Chr(&H6)
'    End If
    
    If mstrCode Like Chr(&H2) & "RE" & Chr(&H3) & "*" Or mstrCode Like Chr(&H2) & "DE" & Chr(&H3) & "*" Then
        mstrCode = Mid(mstrCode, 5)
    End If
    
    If mstrCode Like Chr(&H2) & "DB" & Chr(&H3) & "*" Or mstrCode Like Chr(&H2) & "DB" & Chr(&H3) & "*" Then
        mstrCode = Mid(mstrCode, 5)
    End If
    
    '�ж��Ƿ���һ�����������ݣ����򲻽��н���
    If mstrCode Like Chr(&H2) & "*" & Chr(&H3) Then
        '�ж����������뻹�Ǽ�����
        If mstrCode Like Chr(&H2) & "R*" Then
             'ȡ�����룬��֪ͨ�������˫��ͨѶ
            lngA = UBound(Split(mstrCode, Chr(&H3)))
            For m = 0 To lngA
                str_Tm = Split(mstrCode, Chr(&H3))(m)
                If Len(str_Tm) > 10 Then
                    strRackNo = Mid(str_Tm, 4, 11)
                    str_No = Split(str_Tm, Chr(&H20))(2)
                    str_���� = Trim(Mid(str_Tm, 15))
                    If str_No Like "E*" Then
                        str_Jz = 1
                        str_No = Val(Mid(str_No, 2))
                    End If
                    If str_���� <> "" Then
                        str_No = ""
                    End If
                    strResult = strResult & "||" & Format(Now, "yyyy-mm-dd") & "|" & str_No & "^" & str_Jz & "^" & str_����
                End If
            Next
            strCmd = "1|"
        Else
            '���ն������ݽ��н���
            lngA = UBound(Split(mstrCode, Chr(&H3)))
            For m = 0 To lngA - 1
                str_Hc = Trim(Mid(Split(mstrCode, Chr(&H3))(m), 2))
                If Len(str_Hc) > 10 Then
                    '��ȡ�������
                    str_No = Mid(str_Hc, 10, 4)
                    str_���� = Trim(Mid(str_Hc, 14, 20))
'                    If InStr(str_No, Chr(&H20)) > 0 Then
'                        str_No = Trim(Mid(str_No, 1, InStr(str_No, Chr(&H20))))
'                        str_���� = Trim(Mid(str_No, InStr(str_No, Chr(&H20))))
'                    End If
                    If Mid(str_No, 1, 1) = Chr(&H45) Then
                        str_Jz = 1
                        'ȥ���걾��ǰ��ļ����־
                        str_No = Mid(str_No, 2)
                        str_Hc = Mid(str_Hc, 39)
                    Else
                        str_Jz = 0
                        str_Hc = Mid(str_Hc, 39)
                    End If
                    
                    If str_���� <> "" Then
                        str_No = ""
                    End If
                    strResult = strResult & "||" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "|" & Val(str_No) & "^" & str_Jz & "^" & str_���� & "| |����Ѫ|0"
                    
                    '�Խ�����н���
                    For n = 1 To 100
                        'ÿ������ܳ���Ϊ11�����ڽ�ȡʱȥ�ո��ˣ����Ե��ַ���ʣ�೤��С��9ʱ����������
                        If Len(str_Hc) >= 9 Then
                            str_Item = Mid(str_Hc, 1, 11)
                            str_Hc = Mid(str_Hc, 12)
                            
                            strResult = strResult & "|" & Trim(Mid(str_Item, 1, 3)) & "|" & Val(Trim(Mid(str_Item, 4)))
                        Else
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        strResult = Mid(strResult, 3)
        '������������ջ���
        mstrCode = ""
    End If
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
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    SendSample = ""
    
    blnSuccess = True
    If iSendStep = 0 And strRackNo <> "" Then
        If strResponse Like "*R*" Or strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then '��������
        'Debug.Print "��������"
            If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                '�������
                '�����豸��
                aItem = Split(strSample, "|")
                '�걾��SampleID
                strBuffer = aItem(1)
                strSampleID = aItem(5)
                '��ȡ�����ļܺźͱ���
                If Len(Trim(strSampleID)) > 0 Then
                    'SampleID                          'strData = strData & "|^^" & Mid(strSource, InStr(strSource, "R") + 23, 12)
                    SendSample = "S " & Right(Space(11) & strRackNo, 11) & Right(Space(20) & strSampleID, 20) & Space(4) & "E"
                    'Debug.Print strRackNO
                    
                    'SendSample = "S " & Right(Space(11) & strRackNO, 11) & Right(Space(20) & Mid(strResponse, InStr(strResponse, "R") + 24, 12), 20) & Space(4) & "E"
                Else
                    If strRackNo <> "" Then
                        SendSample = "S " & Right(Space(11) & strRackNo, 11) & Space(24) & "E"
                    Else
                        SendSample = "S         " & Right(String(4, "0") & strBuffer, 4) & Space(24) & "E"
                    End If
                End If
                'Debug.Print SendSample
                For i = 10 To UBound(aItem)
                    SendSample = SendSample & aItem(i)
                Next
                SendSample = SendSample & ""
                'SendSample = SendSample & "SE"
            'Debug.Print SendSample
            
                iSendStep = 0
                strRackNo = ""
            Else
                SendSample = Chr(6)
                iSendStep = iSendStep + 1
            End If
        Else
            SendSample = Chr(6)
            iSendStep = 0
            blnSuccess = False '����ʧ��
        End If
    Else
        SendSample = ""
    End If
End Function

Private Function GetCheckCode(ByVal str_in As String) As String
    '��ȡУ����
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function



