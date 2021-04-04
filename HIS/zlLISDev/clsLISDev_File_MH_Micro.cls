VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MH_Micro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����΢����
'����:����
'����:lwf
'�û�:�ܿ����Ƹ���ҽԺ

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
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strsql As String, rsTmp As New ADODB.Recordset
    Dim str�걾�� As String, str��Ŀ As String
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aLine() As String
    Dim strItem As String
    Dim str_bz As String
    
    
    On Error GoTo errH
    
    Dim rsErs As ADODB.Recordset
    Dim str���� As String
    Dim strData As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .CursorLocation = adUseClient
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb, *.accdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function
        
       strsql = "select a.����ʱ��,b.�걾���,b.ϸ�����,c.ҩ����,c.���,c.MIC,c.���,c.ҩ�����" & _
                " from ������Ϣ a,ϸ����� b,ҩ����� c where a.�걾��� = b.�걾��� and b.�걾���=c.�걾���" & _
                " and Format(a.����ʱ��, 'yyyy/MM/dd HH:mm:ss') >= '" & Format(dtStart, "yyyy/MM/dd HH:mm:ss") & "'" & vbNewLine & _
                " and Format(a.����ʱ��, 'yyyy/MM/dd HH:mm:ss') < '" & Format(dtEnd + 1, "yyyy/MM/dd HH:mm:ss") & "'"

        Set rsTmp = cnAccess.Execute(strsql)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        
        Do While Not rsTmp.EOF
                
            str���� = Format("" & rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss")
            str�걾�� = Val("" & rsTmp("�걾���"))

            str_bz = ""
            
            If rsTmp("ϸ�����") <> "" Then
'                If rsTmp("��ע" & intLoop) <> "" Then str_bz = "|BZ|R^1^" & rsTmp("��ע" & intLoop)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    
                aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |" & _
                    rsTmp("ϸ�����") & "|0"
                If rsTmp("ҩ����") <> "" Then
                    strItem = strItem & "|" & rsTmp("ҩ����") & "|"
                    Select Case Nvl(rsTmp("���"))
                        Case "��ҩ"
                            strItem = strItem & "R"
                        Case "�н�"
                            strItem = strItem & "I"
                        Case "����"
                            strItem = strItem & "S"
                        Case Else
                            strItem = strItem & ""
                    End Select
                    strItem = strItem & "^1^" & rsTmp("MIC")
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & strItem
                    strItem = ""
'                Else
'                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||R^1^" & str_bz
                End If
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    For intLoop = 0 To UBound(aResults)
        Debug.Print aResults(intLoop)
    Next
    
    Exit Function
errH:
    
    MsgBox Err.Description
End Function



'0769-26208501
