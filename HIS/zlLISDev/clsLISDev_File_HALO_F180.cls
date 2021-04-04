VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HALO_F180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:F180��������
'����:
'����:������
'�û�:��������ҽԺ

Public Function GetStartCmd() As String
    '��һ������ʱ����
End Function

Public Function GetEndCmd() As String
    '�ӿڹرյ�ʱ�����
End Function

Public Function GetAnswerCmd() As String
    '��ʱ���͵�ָ��
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '�򿪴�������
    '------------------------------------------------
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

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select format(riqi,'yyyy-mm-dd hh:mm:ss') as ����,BianHao as �걾��,Fenbianyanse as ��ɫ,Xingzhuang as ��״,Hongxibao as ��ϸ��,Baixibao as ��ϸ��,Shiwuxibao as ����ϸ��,Buxiaohuashiwu as ������ʳ��,"
        strSQL = strSQL & "Gouchongluan as ������,Yuanchong as ԭ��,Baonang as ����,Huichongluan as �׳���,Meijunbaozi as ��ĸ��,Qita as ����,Bianchongluan as �޳���,Jushixibao as ����ϸ��,Yinxiehuaxue as ��Ѫ����,"
        strSQL = strSQL & "Qitachongluan as ��������,Jiadishi as �ֵ���,Yinxiemianyi as ŧϸ��,Changdichong as ���γ�,Xialeidengshi as ���׵���,Xianbingdu as �ٲ���,Lunzhuanbingdu as ��״����,Kesaqi as ������,"
        strSQL = strSQL & "Youmenluoxuan as ��������,Zhifangqiu as ֬����,Zhuantiedanbai as ת������,Dangqianzhaopian as ��Ƭ From report where Format(riqi,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
        strSQL = strSQL & " And Format(riqi,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' and wancheng = 1" & _
                 IIf(lngSampleNO = -1, "", " And val(bianhao)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = rsTmp("����") & "|" & rsTmp("�걾��") & "| |���|0|FBYS|" & Trim(rsTmp("��ɫ")) & "|FBXZ|" & Trim(rsTmp("��״")) & "|HXB|" & Trim(rsTmp("��ϸ��")) & _
                     "|BXB|" & Trim(rsTmp("��ϸ��")) & "|SWXB|" & Trim(rsTmp("����ϸ��")) & "|BXHSW|" & Trim(rsTmp("������ʳ��")) & "|GCL|" & Trim(rsTmp("������")) & "|YC|" & Trim(rsTmp("ԭ��")) & _
                     "|BN|" & Trim(rsTmp("����")) & "|HCL|" & Trim(rsTmp("�׳���")) & "|JMJ|" & Trim(rsTmp("��ĸ��")) & "|QT|" & IIf(Trim(rsTmp("��������")) = "", "δ���", Trim(rsTmp("����"))) & "|BCL|" & Trim(rsTmp("�޳���")) & _
                     "|TSXB|" & Trim(rsTmp("����ϸ��")) & "|YXSY|" & Trim(rsTmp("��Ѫ����")) & "|QTCL|" & IIf(Trim(rsTmp("��������")) = "", "δ���", Trim(rsTmp("��������"))) & "|JDS|" & Trim(rsTmp("�ֵ���")) & "|NXB|" & Trim(rsTmp("ŧϸ��")) & _
                     "|CDC|" & Trim(rsTmp("���γ�")) & "|XLDS|" & Trim(rsTmp("���׵���")) & "|XBD|" & Trim(rsTmp("�ٲ���")) & "|LZBD|" & Trim(rsTmp("��״����")) & "|KSQ|" & Trim(rsTmp("������")) & _
                     "|YMLX|" & Trim(rsTmp("��������")) & "|ZFQ|" & Trim(rsTmp("֬����")) & "|ZTDB|" & rsTmp("ת������")
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    '������ɣ��ر��ļ�
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults

    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_HALO_F180"
    Close #lngErrFileNo
End Function