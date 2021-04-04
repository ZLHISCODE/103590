VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_DSI905"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�����
'�����ͺţ�DSI905
'�û����������ο����
'��ս�У�������
'���ڣ�2014-09-11
'
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
    Dim J As Integer
    Dim lngID As String
    Dim strLine As String
    Dim Var_date As Variant
    Dim var_date1 As Variant

    mstrCode = mstrCode & strSource

    lng_S = InStr(mstrCode, "1")
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, "*" & Chr$(&H1A))
    Else
        lng_E = InStr(mstrCode, "*" & Chr$(&H1A))
    End If

    Do While lng_S > 0 And lng_E > 0

        strCode = Mid$(mstrCode, lng_S, lng_E - lng_S)
        mstrCode = Mid$(mstrCode, lng_E)

        '��ȡ�걾��
        strCode = Replace$(strCode, "*", " ")
        Var_date = Split(strCode, " ")

        str_Day = Format$(Now, "yyyy/mm/dd hh:mm:ss")

        str_No = Right$(Var_date(0), 3)
        If UBound(Var_date) = 6 Then
            strLine = "|K|" & Val(Var_date(1)) / 100 & "|Na|" & Val(Var_date(2)) / 10 & "|Cl|" & Val(Var_date(3)) / 10
            strLine = strLine & "|Ga|" & Val(Var_date(4)) / 100 & "|nGa|" & Val(Var_date(5)) / 100 & "|PH|" & Val(Var_date(6)) / 100
        Else
            strLine = "|K|" & Val(Var_date(1)) / 100 & "|Na|" & Val(Var_date(2)) / 10 & "|Cl|" & Val(Var_date(4)) / 10
            strLine = strLine & "|Ga|" & Val(Var_date(5)) / 100 & "|nGa|" & Val(Var_date(6)) / 100 & "|PH|" & Val(Var_date(7)) / 100
        End If
        str��� = str��� & "||" & str_Day & "|" & str_No & "| |ѪҺ|0" & strLine

        lng_S = InStr(mstrCode, "1")
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, "*" & Chr$(&H1A))
        Else
            lng_E = InStr(mstrCode, "*" & Chr$(&H1A))
        End If
    Loop

    strResult = Mid$(str���, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
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