VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Ruby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Ѫ����
'�����ͺţ�Ruby
'���ң���������
'�û������ǹſ�ҽԺ
'������
'���ڣ�2014-11-28
Private mstrCode As String    '��������
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


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


'----------------------�ı���ʽ�ӿں���---------------------------------------------------
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '��׼���������ļ���ȡ������
    '   strFile������·�����ļ���
    '   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    '   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    '   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
    '   ���صļ���������(������������밴���±�׼��֯���)
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������

End Function


'----------------------------------------------------------------------------------------

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
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
    Dim lng_S As Long, lng_E As Long
    Dim str_Sample As String, str_Day As String
    Dim str_No As String, str_Item As String, str_Value As String
    Dim INR_Value As String
    Dim str��� As String
    Dim strCode As String
    Dim j As Integer
    Dim lngID As String
    Dim strLine As String
    Dim Var_date As Variant
    Dim str��Ŀ As String
    Dim strVLineLeft  As String '���ֽ���
    Dim strVLineRight As String  '�Ҳ�ֽ���
    Dim strImageName  As String   'ͼ������
    Dim strImage  As String       'ͼ������
    '    mstrCode = ""
    mstrCode = mstrCode & strSource

    'Ӧ��
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If

    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
    Else
        lng_E = InStr(mstrCode, "L|1|N")

    End If

    Do While lng_S > 0 And lng_E > 0

        strLine = ""
        strCode = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E)

        Var_date = Split(strCode, Chr(&HD))
        For j = LBound(Var_date) To UBound(Var_date)
            strLine = Var_date(j)
            strLine = strLine

            If InStr(strLine, "No_ID") > 0 Then

                str_Day = "": str_No = "":  str��Ŀ = ""
                str_Day = Format(Now, "yyyy-mm-dd hh:mm:ss")

                '�걾��
                str_No = Split(strLine, "|")(3)

            ElseIf InStr(strLine, "R|") > 0 And InStr(strLine, "F|") > 0 Then
                '����
                If str_Day = "" Then
                    str_Day = Split(strLine, "|")(11)
                    str_Day = Mid(str_Day, 1, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Mid(str_Day, 7, 2) & "��" & _
                              Mid(str_Day, 9, 2) & ":" & Mid(str_Day, 11, 2) & ":" & Mid(str_Day, 13, 2)
                    str_Day = Format(str_Day, "yyyy-mm-dd hh:mm:ss")
                End If

                str_Item = Split(strLine, "|")(2)
                str_Item = Split(str_Item, "^")(6)

                str_Value = Split(strLine, "|")(3)

                If str_Value = "" Then
                    str_Value = 0
                End If

                str��Ŀ = str��Ŀ & "|" & str_Item & "|" & str_Value
            ElseIf Split(strLine, "|")(0) = "M" Then
                'ֱ��ͼ
                If Split(strLine, "|")(3) = "WB1" Or Split(strLine, "|")(3) = "RBC" Or Split(strLine, "|")(3) = "PLT" Then
                    strVLineLeft = Split(Split(strLine, "|")(5), "\")(0)
                    strVLineRight = Split(Split(strLine, "|")(5), "\")(1)
                    strImageName = IIf(Split(strLine, "|")(3) = "WB1", "WBC", Split(strLine, "|")(3))
                    strLine = Split(strLine, "|")(6)
                    strImage = strImage & "^" & strImageName & ";5;255,255;10,50,50,10;0-0,50-50,100-100,150-150,200-200,250-250|0-0,50-50,100-100,150-150,200-200,250-250;" & _
                               Base64Decode(strLine) & ";" & strVLineLeft & "," & strVLineRight
                End If
            End If

        Next

        str��� = str��� & "||" & str_Day & "|" & str_No & "|������|ѪҺ|0" & str��Ŀ & vbCrLf & Mid(strImage, 2)

        lng_S = InStr(mstrCode, Chr(&HD))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
        Else
            lng_E = InStr(mstrCode, "L|1|N")
        End If
    Loop

    strResult = Mid(str���, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_Ruby"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
End Function



'---------------------------ø���ǽӿ��Զ���ָ��--------------------
Public Function CmdAnalyse(strCmdType As String) As String

End Function


Public Function StrValue(strItem As String, strC As Integer) As String

Dim str_j As String
str_j = strItem

StrValue = Split(strItem, "^")(strC)

End Function



Private Function Base64Decode(strEncoded As String) As String '����

    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT, bRet() As Byte
    Dim i As Long, j As Long
    Dim strResult   As String
    Dim var()   As String
'    Dim BASE64CHR As String
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    Dim bRetTemp(3) As Byte
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63
        
        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
        bRet((i \ 4) * 3) = bT \ 65536
        If (i \ 4) * 3 Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3)
        End If
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        If ((i \ 4) * 3 + 1) Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3 + 1)
        End If
        bRet((i \ 4) * 3 + 2) = bT And 255
        If ((i \ 4) * 3 + 2) Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3 + 2)
        End If
    Next
    Base64Decode = strResult
End Function