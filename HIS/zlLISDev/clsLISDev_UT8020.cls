VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UT8020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:URIT-8020����������
'����:������
'����:�µ���
'�û�:�ɶ������Բ�ҽԺ

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
    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String, mstrCode As String
    Dim str_�ʿ� As String, str_Return As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    
    strResult = "": strReserved = ""
    
    lng_S = InStr(strSource, Chr(&H7E))
    lng_E = InStr(strSource, Chr(&H40))
    
    If lng_S > 0 And lng_E > 0 Then
        str_Code = Mid(strSource, lng_S, lng_E)
        strReserved = Mid(strSource, lng_E + 1)
        
        str_Return = "": str_�걾�� = "": str_���� = ""
        var_Code = Split(str_Code, Chr(&HA))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If lng_count = 0 Then
                '�걾�ţ�����
                var_Filed = Split(str_Segment, "`")
                str_�걾�� = var_Filed(5)
                str_���� = Trim(var_Filed(11))
                str_���� = Replace(str_����, Chr(&HD), "")
                str_���� = Trim(str_����) & " " & Format(Now, "HH:MM:SS")
                If Not IsDate(str_����) Then str_���� = Format(Now, "YYYY-MM-DD HH:MM:SS")
            Else
               If str_Segment <> "@" Then
                  var_Filed = Split(str_Segment, "`")
                  str_��Ŀ = Trim(var_Filed(1))
                  str_��� = Trim(var_Filed(3))
                  str_Return = str_Return & "|" & str_��Ŀ & "|" & str_���
               End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
        End If
        If strResult <> "" Then strResult = Mid(strResult, 3)
    Else
      strReserved = strSource
    End If
End Sub
