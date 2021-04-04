VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_K_Lite6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�����
'�����ͺţ�K_Lite6
'�û���ɽ��ʡ����ɽҽԺ
'��ս�У�������
'���ڣ�2014-07-09
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
    Dim str_Day As String
    Dim str_No As String, Str_Item As String, str_Value As String
  
    Dim strResultTmp As String
    Dim strCode As String
    Dim j As Integer
    Dim strLine As String
    Dim var_Date As Variant
    
    mstrCode = mstrCode & strSource
    
    On Error GoTo errHand
    
    lng_S = InStr(mstrCode, Chr(&HA))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, Chr(&HD))
    Else
        lng_E = InStr(mstrCode, Chr(&HD))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        
        strCode = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E)
        '��ȡ�걾��
        var_Date = Split(strCode, Chr(&HA))
       
     
        For j = LBound(var_Date) To 10
        
            strLine = var_Date(j)
            
            If IsDate(strLine) Then
                str_Day = Format(strLine, "yyyy-mm-dd hh:mm")
                j = 2
            ElseIf InStr(strLine, "PID") > 0 Then
                str_No = Mid(strLine, 6)
                strResultTmp = strResultTmp & "||" & str_Day & "|" & str_No & "|������|ѪҺ|0"
                
            ElseIf InStr(strLine, "PID") = 0 And j > 3 And strLine <> "" Then
                Str_Item = Trim(Mid(strLine, 1, 3))
                              
                str_Value = Trim(Mid(strLine, 6, 7))
                
                If str_Value = "" Then
                    str_Value = "0"
                End If
                strResultTmp = strResultTmp & "|" & Str_Item & "|" & str_Value
                    
            End If
        Next
        
        lng_S = InStr(mstrCode, Chr(&HA))
        
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, Chr(&HD))
        Else
            lng_E = InStr(mstrCode, Chr(&HD))
        End If
        
    Loop
    strResult = Mid(strResultTmp, 3)
    Exit Sub
errHand:
    WriteLog "clsLISdev_K_Lite6", Err.Description, "�����������" & mstrCode
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





