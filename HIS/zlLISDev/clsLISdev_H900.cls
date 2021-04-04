VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_H900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�����
'�����ͺţ�H900
'�û���������ϣ����ҽ�õ������޹�˾
'��ս�У�������
'���ڣ�2014-10-14
'ҽԺ���ӽ�����ҽԺ
Private mstrCode As String    '��������

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
    On Error GoTo ErrHand '�򿪴�������
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
    Dim var_date1 As Variant

    '    mstrCode = ""
    mstrCode = mstrCode & strSource

    lng_S = InStr(mstrCode, "Sample")
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, "End")
    Else
        lng_E = InStr(mstrCode, "End")
    End If

    Do While lng_S > 0 And lng_E > 0

        strCode = Mid(mstrCode, lng_S, lng_E - lng_S)
        mstrCode = Mid(mstrCode, lng_E)
        '��ȡ�걾��

        Var_date = Split(strCode, " ")

        str_Day = Format(Trim(Mid(strCode, InStr(strCode, "Date") + 4, 10)), "yyyy/mm/dd")
        str_Day = str_Day & " " & Trim(Mid(strCode, InStr(strCode, "Time") + 4, 10))
        str_Day = Format(str_Day, "yyyy/mm/dd hh:mm:ss")

        str_No = Trim(Mid(strCode, InStr(strCode, "PAT") + 3, 10))

        strLine = "|K|" & Trim(Mid(strCode, InStr(strCode, "K") + 1, 6)) & "|Na|" & Trim(Mid(strCode, InStr(strCode, "Na") + 2, 6)) & "|Cl|" & Trim(Mid(strCode, InStr(strCode, "Cl") + 2, 6))
        strLine = strLine & "|iCa|" & Trim(Mid(strCode, InStr(strCode, "iCa") + 3, 6)) & "|TCa|" & Trim(Mid(strCode, InStr(strCode, "TCa") + 3, 6)) & "|nCa|" & Trim(Mid(strCode, InStr(strCode, "nCa") + 3, 6))
        strLine = strLine & "|TCO2|" & Trim(Mid(strCode, InStr(strCode, "TCO2") + 4, 10)) & "|pH|" & Trim(Mid(strCode, InStr(strCode, "pH") + 2, 6)) & "|AG|" & Trim(Mid(strCode, InStr(strCode, "AG") + 2, 10))
        strLine = strLine & "|Mg|" & Trim(Mid(strCode, InStr(strCode, "Mg") + 2, 10)) & "|Li|" & Trim(Mid(strCode, InStr(strCode, "Li") + 2, 10))

        str��� = str��� & "||" & str_Day & "|" & str_No & "|������|ѪҺ|0" & strLine

        lng_S = InStr(mstrCode, "Sample")
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, "End")
        Else
            lng_E = InStr(mstrCode, "End")
        End If
    Loop

    strResult = Mid(str���, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_GFM3000"
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

Public Function StrValue(strN As String) As String

If InStr(strN, "(") > 0 Then
    StrValue = Split(strN, "(")(0)
Else
    StrValue = strN
End If


End Function