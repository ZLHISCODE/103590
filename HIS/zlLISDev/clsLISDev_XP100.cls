VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XP100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:XP100Ѫ������
'����:
'����:�����
'�û�:��Ԫ��������ҽԺ
'�޸ģ����� ����ASTM��ʽ����

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String, str_�ʿ� As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String, iRows As Integer
    Dim varLine As Variant, blnEnd As Boolean
    Dim i As Integer, str_Line As String, iRow As Integer
    Dim lngLength As Long, strData As Variant
    Dim strRequest As String
    Dim strLine As String, strVarry() As String
    
    '��ʼ������ֵ
'    If strSource <> "" Then strReserved = strReserved & strSource
On Error GoTo errh
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    Call writelislog(" �յ���" & strSource)
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '�յ���ʼ�� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&HA) & Chr(&HD)) > 0 Then                              '�յ��н����� 03 ,Ӧ�� 06
         strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H3)) > 0 Then                                         '�յ���ʼ�� 03 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H4)) > 0 Then                                         '�յ������� 04 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
'    lng_S = InStr(strReserved, Chr(&H2))
'    lng_E = InStr(strReserved, "")
    Do While lng_S > 0 And lng_E > 0
        If lng_S <= lng_E Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            str_��Ŀ = ""
'            If str_Data Like "*/*/*" Then
'                str_�ʿ� = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
'                str_ʱ�� = Format(Trim(Mid(str_Data, 2, 8)), "yyyy-MM-dd") & " " & Format(Trim(Mid(str_Data, 10, 5)), "HH:mm")
'                str_������ = Val(Mid(str_Data, 25, 3))
'                str_��Ŀ = str_��Ŀ & "|" & "WBC" & "|" & Val(Trim(Mid(str_Data, 83, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "RBC" & "|" & Val(Trim(Mid(str_Data, 88, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "HGB" & "|" & Val(Trim(Mid(str_Data, 93, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "HCT" & "|" & Val(Trim(Mid(str_Data, 98, 5))) * (10 ^ 2)
'                str_��Ŀ = str_��Ŀ & "|" & "MCV" & "|" & Val(Trim(Mid(str_Data, 103, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "MCH" & "|" & Val(Trim(Mid(str_Data, 108, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "MCHC" & "|" & Val(Trim(Mid(str_Data, 113, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "PLT" & "|" & Val(Trim(Mid(str_Data, 118, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "LYM%" & "|" & Val(Trim(Mid(str_Data, 123, 5))) * (10 ^ 2)
'                str_��Ŀ = str_��Ŀ & "|" & "MDX%" & "|" & Val(Trim(Mid(str_Data, 128, 5))) * (10 ^ 2)
'                str_��Ŀ = str_��Ŀ & "|" & "NEUT%" & "|" & Val(Trim(Mid(str_Data, 133, 5))) * (10 ^ 2)
'                str_��Ŀ = str_��Ŀ & "|" & "LYM#" & "|" & Val(Trim(Mid(str_Data, 138, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "MXD#" & "|" & Val(Trim(Mid(str_Data, 143, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "NEUT#" & "|" & Val(Trim(Mid(str_Data, 148, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "RDW-C" & "|" & Val(Trim(Mid(str_Data, 153, 5))) * (10 ^ 2)
'                str_��Ŀ = str_��Ŀ & "|" & "RDW-S" & "|" & Val(Trim(Mid(str_Data, 158, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "PDW" & "|" & Val(Trim(Mid(str_Data, 163, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "MPV" & "|" & Val(Trim(Mid(str_Data, 168, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "P-LCR" & "|" & Val(Trim(Mid(str_Data, 173, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "1" & "|" & Val(Trim(Mid(str_Data, 178, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "2" & "|" & Val(Trim(Mid(str_Data, 183, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "3" & "|" & Val(Trim(Mid(str_Data, 188, 5)))
'                If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
'                    strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |" & "Ѫ��|" & str_�ʿ� & str_��Ŀ
'                    Debug.Print strResult
'                End If
'            End If
            If str_Data Like "*1H|*" Then 'XP100����ASTM��ʽ����
                strVarry = Split(str_Data, Chr(&HD))
                For i = LBound(strVarry) To UBound(strVarry)
                    strLine = strVarry(i)
                    If InStr(strLine, "R|") > 0 Then '��ȡ��������
                        str_��Ŀ = str_��Ŀ & "|" & Split(Split(strLine, "|")(2), "^")(4) & "|" & Val(Split(strLine, "|")(3))
                        str_ʱ�� = Split(strLine, "|")(12)
                        str_ʱ�� = Format(Left(str_ʱ��, 4) & "-" & Mid(str_ʱ��, 5, 2) & "-" & Mid(str_ʱ��, 7, 2) & " " & Mid(str_ʱ��, 9, 2) & ":" & Mid(str_ʱ��, 11, 2) & ":" & Right(str_ʱ��, 2), "yyyy-mm-dd HH:mm:ss")
                    End If
                    If InStr(strLine, "O|") > 0 Then '��ȡ�걾��
                        str_������ = Val(Split(Split(strLine, "|")(3), "^")(2))
                    End If
                Next
                If str_������ <> "" And str_��Ŀ <> "" Then
                    strResult = str_ʱ�� & "|" & str_������ & "| |ѪҺ|0" & str_��Ŀ
'                    Call writelislog(" �ӿڷ��أ�" & strResult)
                End If
            End If
            If str_Data Like "/*/*" Then
                str_�ʿ� = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
                str_ʱ�� = Format(Trim(Mid(str_Data, 2, 8)), "yyyy-MM-dd") & " " & Format(Trim(Mid(str_Data, 10, 5)), "HH:mm")
                str_������ = Val(Mid(str_Data, 25, 3))
                strRequest = Mid$(str_Data, InStr(str_Data, "") + 2)
                str_��Ŀ = str_��Ŀ & "|" & "WBC" & "|" & Val(Trim(Mid(strRequest, 1, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "RBC" & "|" & Val(Trim(Mid(strRequest, 6, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "HGB" & "|" & Val(Trim(Mid(strRequest, 11, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "HCT" & "|" & Val(Trim(Mid(strRequest, 16, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "MCV" & "|" & Val(Trim(Mid(strRequest, 21, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "MCH" & "|" & Val(Trim(Mid(strRequest, 26, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "MCHC" & "|" & Val(Trim(Mid(strRequest, 31, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "PLT" & "|" & Val(Trim(Mid(strRequest, 36, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "LYM%" & "|" & Val(Trim(Mid(strRequest, 43, 4)))
                str_��Ŀ = str_��Ŀ & "|" & "MDX%" & "|" & Val(Trim(Mid(strRequest, 48, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "NEUT%" & "|" & Val(Trim(Mid(strRequest, 53, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "LYM#" & "|" & Val(Trim(Mid(strRequest, 58, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "MXD#" & "|" & Val(Trim(Mid(strRequest, 63, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "NEUT#" & "|" & Val(Trim(Mid(strRequest, 67, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "RDW-C" & "|" & Val(Trim(Mid(strRequest, 72, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "RDW-S" & "|" & Val(Trim(Mid(strRequest, 77, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "PDW" & "|" & Val(Trim(Mid(strRequest, 82, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "MPV" & "|" & Val(Trim(Mid(strRequest, 87, 5)))
                str_��Ŀ = str_��Ŀ & "|" & "P-LCR" & "|" & Val(Trim(Mid(strRequest, 91, 5)))
                If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
                    strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |" & "Ѫ��|" & str_�ʿ� & str_��Ŀ
                    Debug.Print strResult
                End If
            End If
            If str_Data Like "*D1U*" Then
                str_�ʿ� = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
                str_ʱ�� = Mid(str_Data, 45, 8)
                str_ʱ�� = Format(Left(str_ʱ��, 4) & "-" & Mid(str_ʱ��, 5, 2) & "-" & Mid(str_ʱ��, 7, 2), "yyyy-mm-dd")
                str_������ = Val(Mid(str_Data, 25, 3))
                strRequest = Mid$(str_Data, 68)
                str_��Ŀ = str_��Ŀ & "|" & "WBC" & "|" & Val(Trim(Mid(strRequest, 10, 4))) / 10 ^ 2
                str_��Ŀ = str_��Ŀ & "|" & "RBC" & "|" & Val(Trim(Mid(strRequest, 15, 4))) / 10 ^ 3
                str_��Ŀ = str_��Ŀ & "|" & "HGB" & "|" & Val(Trim(Mid(strRequest, 19, 5))) / 10
                str_��Ŀ = str_��Ŀ & "|" & "HCT" & "|" & Val(Trim(Mid(strRequest, 24, 4))) / 10 ^ 3
                str_��Ŀ = str_��Ŀ & "|" & "MCV" & "|" & Val(Trim(Mid(strRequest, 29, 4))) / 10
                str_��Ŀ = str_��Ŀ & "|" & "MCH" & "|" & Val(Trim(Mid(strRequest, 33, 5))) / 10
                str_��Ŀ = str_��Ŀ & "|" & "MCHC" & "|" & Val(Trim(Mid(strRequest, 39, 4)))
                str_��Ŀ = str_��Ŀ & "|" & "PLT" & "|" & Val(Trim(Mid(strRequest, 44, 4)))
                str_��Ŀ = str_��Ŀ & "|" & "LYM%" & "|" & Val(Trim(Mid(strRequest, 49, 4))) / 10 ^ 4
                str_��Ŀ = str_��Ŀ & "|" & "MDX%" & "|" & Val(Trim(Mid(strRequest, 54, 4))) / 10 ^ 3
                str_��Ŀ = str_��Ŀ & "|" & "NEUT%" & "|" & Val(Trim(Mid(strRequest, 59, 4))) / 10 ^ 3
                str_��Ŀ = str_��Ŀ & "|" & "LYM#" & "|" & Val(Trim(Mid(strRequest, 64, 5))) / 10 ^ 2
                str_��Ŀ = str_��Ŀ & "|" & "MXD#" & "|" & Val(Trim(Mid(strRequest, 69, 5))) / 10 ^ 2
                str_��Ŀ = str_��Ŀ & "|" & "NEUT#" & "|" & Val(Trim(Mid(strRequest, 74, 4))) / 10
                str_��Ŀ = str_��Ŀ & "|" & "RDW-C" & "|" & Val(Trim(Mid(strRequest, 84, 5))) / 10 ^ 4
                str_��Ŀ = str_��Ŀ & "|" & "RDW-S" & "|" & Val(Trim(Mid(strRequest, 79, 4))) / 10
'                str_��Ŀ = str_��Ŀ & "|" & "PDW" & "|" & Val(Trim(Mid(strRequest, 82, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "MPV" & "|" & Val(Trim(Mid(strRequest, 87, 5)))
'                str_��Ŀ = str_��Ŀ & "|" & "P-LCR" & "|" & Val(Trim(Mid(strRequest, 91, 5)))
                If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
                    strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |" & "Ѫ��|" & str_�ʿ� & str_��Ŀ
                    Debug.Print strResult
                End If
            End If
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = InStr(mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errh:
    WriteLog "XP100", "�������", Err.Description
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

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_XP100_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function





