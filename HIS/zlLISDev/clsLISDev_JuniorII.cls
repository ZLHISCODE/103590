VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_JuniorII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��������� Junior II
'����:
'����:�¶�
'�û�:����ҽԺ
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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim varIn As Variant '�����������
    Dim StrTime As String
    
    Dim lng_S As Long, lng_E As Long
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String
    Dim str��ʽ As String, strItem As String
    Dim strCode As String
    On Error GoTo errHandle
    str�����˱걾�ʿ� = " |��Һ|0"
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource = "" Then Exit Sub
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    If Not (lng_S > 0 And lng_E > 0) Then
        lng_S = InStr(mstrCode, "URINALYSIS")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&HC))
        Else
            lng_E = InStr(mstrCode, Chr(&HC))
        End If
    End If
    
    If lng_S > 0 And lng_E > 0 Then
        strCode = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        varIn = Split(strCode, vbNewLine)
        strLine = ""
        For iStrPoint = LBound(varIn) To UBound(varIn)
        
            strLine = varIn(iStrPoint)
            
            '--- �걾��"               Seq.no. 12"
            If InStr(strLine, "Seq.no.") > 0 Then
                str�걾�� = Val(Trim(Mid(strLine, InStr(strLine, ".no.") + 4)))
                If str�걾�� = "0" Then Exit Sub
            End If
            
            '--- ����  "07.03.20           10:54"
            
            str��ʽ = Trim(ReadIni("SET", "���ڸ�ʽ", App.Path & "\JuniorII.ini"))
            If str��ʽ = "" Then
                WriteIni "SET", "���ڸ�ʽ", "��.��.��", App.Path & "\JuniorII.ini"
                str��ʽ = "��.��.��"
            End If
            
            If strLine Like "*.*.*:*" And InStr(strLine, "Seq.no.") <= 0 Then
                strLine = Mid(strLine, 1, 10)
                If str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(2)), "yyyy-MM-dd")
                ElseIf str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1)), "yyyy-MM-dd")
                ElseIf str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(0)), "yyyy-MM-dd")
                End If
            ElseIf strLine Like "*Date:*.*.*" Then
                strLine = Split(strLine, "Date:")(1)
                If str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(2)), "yyyy-MM-dd")
                ElseIf str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1)), "yyyy-MM-dd")
                ElseIf str��ʽ = "��.��.��" Then
                    str���� = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(0)), "yyyy-MM-dd")
                End If
            End If
            
            '-- ��ɫ
            If InStr(strLine, "Color:") > 0 And strResult <> "" Then
                strResult = strResult & "|Color|" & Trim(Mid(strLine, 11))
                Exit Sub
            End If
            
            '-- ָ�� "  SG :       1.015 "
            strָ�� = ""
            If InStr(strLine, ":") = 6 Then
                strָ�� = strLine
            End If
            
            If strResult = "" And str�걾�� <> "" And str���� <> "" And strָ�� = "" Then
                strResult = str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ�
            ElseIf strResult <> "" And str�걾�� <> "" And str���� <> "" And strָ�� <> "" Then
                strItem = IIf(Trim(Mid(strָ��, 7, 12)) = "", "����", Trim(Mid(strָ��, 7, 12)))
                strItem = Replace(strItem, "/ul", "")
                strItem = Replace(strItem, "mg/dl", "")
                strResult = strResult & "|" & Trim(Mid(strָ��, 2, 4)) & "|" & Trim(strItem)
            End If
        Next
    End If
    
    Exit Sub
errHandle:
    WriteLog "JuniorII", Err.Description, ""

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

