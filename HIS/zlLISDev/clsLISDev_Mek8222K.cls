VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek8222K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Mek 822k
'����:�ձ����
'����:����ǿ
'�û�:��ϼҽԺrs-232��������
'-- ����ͼ������ʱ�ĸ�ʽ:
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ  2=Ѫ����ճ����������  3=Ѫ������  100=ͼƬ����
    '                     0) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);Y1;Y2;Y3;Y4;Y5...
    '                     1) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    '                     2) ճ����������:ͼ������;ͼ�񻭷�;��������;���߼��������;�������������
    '                                   ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
    '                                   ���߼��������:ճ������1�ĸߵ�͵͵�����|ճ������2�ĸߵ�͵͵�����~���е�����,���е�����,���е�����
    '                                   �������������:Y�����������,X����,Y����~X�����������,X����,Y����
    '                        ��:ճ����������;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;VIS(mPa.s),25,20~SHR(1/S),195,1
    '                     3) Ѫ������:ͼ������;ͼ�񻭷�;��������;�������;�������������
    '                                   ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
    '                                   �������:Ѫ��ֵ1,Ѫ��ֵ2,....Ѫ��ֵ30
    '                                   �������������:Y�����������,X����,Y����~X�����������,X����,Y����
    '                        ��:Ѫ������;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;.5,.5,1,1,1,1.5,1.5,2,2,2,2.5,3,3,3.5,4,4.5,5.5,6.5,8,9,10.5,11.5,12.5,13.5,14.5,15.5,16.5,18,19,20;Ѫ��ֵ(mm),5,36~ʱ��(m),55,1
    '                     4) PLTͼ��ͼ������;ͼ�񻭷�;��������;�������
    '                               ���� �������ݣ�Y����,X����,X����-X������ʾ������,....
    '                                    �������: Y1,Y2,Y3,......|Y1,Y2,Y3,......
    '                        ��:PLT;4;200,262;0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,4,4,7,7,12,12,17,17,20,20,25,25,30,30,33,33,36,36,41,41,43,43,44,44,46,46,47,47,47,47,47,47,46,46,46,46,44,44,44,44,43,43,41,41,39,39,38,38,36,36,35,35,33,33,31,31,30,30,28,28,27,27,25,25,23,23,22,22,22,22,20,20,19,19,17,17,15,15,15,15,14,14,12,12,12,12,11,11,11,11,9,9,9,9,9,9,7,7,7,7,7,7,6,6,6,6,6,6,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,3,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,7,7,9,9,8,8,9,9,12,12,16,16,22,22,26,26,30,30,35,35,36,36,37,37,39,39,42,42,44,44,46,46,46,46,44,44,43,43,40,40,37,37,37,37,37,37,39,39,37,37,36,36,32,32,29,29,25,25,23,23,22,22,22,22,21,21,19,19,18,18,16,16,16,16,15,15,15,15,15,15,14,14,12,12,11,11,9,9,9,9,8,8,8,8,7,7,7,7,7,7,7,7,7,7,8,8,7,7,7,7,5,5,4,4,4,4,2,2,4,4,4,4,2,2,2,2,4,4
    '                   100) ͼƬ����:ͼ������;ͼ�񻭷�;[��ȡ���ݺ��Ƿ�ɾ��];ȫ·��
    '                        ��:WBC Fsc;100;1;C:\tempfile.gif
    


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

    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim strImage As String  'ͼ������
    Dim varImage As Variant 'ͼ������
    Dim intItem As Integer
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*MEK-8222**" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            strLine = Mid(strReserved, InStr(strReserved, "MEK-8222"), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = "": intItem = 0
            '����
            varLine = Split(strLine, Chr(13))

            '�������
            For iRow = 0 To UBound(varLine)
                Select Case iRow
                    Case 15, 16, 17 '����
                    If iRow = 15 Then
                        str���� = varLine(iRow) & "-"
                    ElseIf iRow = 16 Then
                        str���� = str���� & varLine(iRow) & "-"
                    Else
                        str���� = Format(str���� & varLine(iRow), "yyyy-MM-dd")
                    End If
                    Case 22 '�걾��
                    str�걾�� = Val(varLine(iRow))
                    str�����˱걾�ʿ� = " |ѪҺ|0"
                    Case 23 To 44
                    intItem = intItem + 1
                    strTmp = Replace(Replace(Trim(varLine(iRow)), "L", ""), "H", "")

                    strItem = strItem & "|" & intItem & "|" & Val(strTmp)
                End Select
            Next

            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            '            If Not strReserved Like "*MEK-8222**" Then
            blnEnd = True
            '            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

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

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function