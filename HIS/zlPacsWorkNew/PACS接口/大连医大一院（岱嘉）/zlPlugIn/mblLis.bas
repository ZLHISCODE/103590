Attribute VB_Name = "mblLis"
Public glngCount As Long
Public gblnInit As Boolean
Public gblnDebug As Boolean

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW As Long = 5
Public gFso As New FileSystemObject


Public Sub WriteLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    strFileName = App.Path & "\zlPlugIn_" & Format(Date, "yyyyMMdd") & ".LOG"
    If Not gFso.FileExists(strFileName) Then Call gFso.CreateTextFile(strFileName)
    Set objStream = gFso.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strDate & ":" & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub

 

