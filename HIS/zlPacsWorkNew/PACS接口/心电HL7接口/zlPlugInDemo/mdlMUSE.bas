Attribute VB_Name = "mdlMUSE"
Option Explicit

''''''''���˵��''''''''''''''''''''''
'''˵����''''''''''''''''''''''''''''''''''''''''''
'''1�������ӳ����У�����MUSE�ĵ�ϵͳ������Ĳ�����Ҫ��mdlMUSEģ����ʵ�֡�
'''2��ͨ������MUSE����ֱ�ӵ��ñ�������������ĵ����������ӡ�


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''˵�����˴�����MUSE�Ĺ��ܣ������Ӧ�Ĺ�����������������ģ��clsPlugIn�еĵ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gstrFunc_MUSE�ĵ������� = "�ĵ�������"

Public Function ShowMUSEViewer(ByVal varKeyId As Variant) As Boolean
'˵������ʾMUSE�������

'������ varKeyId --- ҽ��ID

    Dim blnErr As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strURL As String
    Dim i As Integer
    
    On Error GoTo err
    
    ShowMUSEViewer = False
    
    '��HIS�����ݿ��в��ұ���ҽ��ID��Ӧ���ĵ�ϵͳ����URL
    strSQL = "Select ִ��˵�� From ����ҽ������ Where ҽ��ID = " & varKeyId
    Set rsTemp = gcnOracle.Execute(strSQL)
    
    '��Ϊֻ֪��ҽ��ID����֪�����ͺţ����ڳ�������Ҫѭ�����ҵ�һ����ִ��˵���ļ�¼���������ĵ�ϵͳ�ļ����
    For i = 1 To rsTemp.RecordCount
        strURL = IIf(IsNull(rsTemp!ִ��˵��), "", rsTemp!ִ��˵��)
        If strURL <> "" Then
            Exit For
        End If
        rsTemp.MoveNext
    Next i
    
    If strURL <> "" Then
        '�������
        Shell "explorer " & strURL, 0
        ShowMUSEViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "MUSE�ӿڴ���"
    err.Clear
End Function
