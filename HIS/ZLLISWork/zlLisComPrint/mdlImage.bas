Attribute VB_Name = "mdlImage"
Option Explicit
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Private mobjImg As Object                            'zllisdev.dll����

Public Function ReadSampleImage(lngSampleID As Long, strChar() As String, Optional strErr As String) As Boolean
    '����   ����걾��ͼ�񷵻ض���������
    '��ͼ��
    Dim rsImage As ADODB.Recordset
    Dim intLoop As Integer, strReturn As String
    Dim varTmp As Variant, strDir As String
    Dim i As Integer
    
    
    On Error GoTo errH
    
    
    strErr = ""
    strDir = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strDir) Then Call gobjFSO.CreateFolder(strDir)
    
    If mobjImg Is Nothing Then
        Set mobjImg = CreateObject("zlLisDev.clsDrawGraph")
        Call mobjImg.GetSampleImgInit(gSysInfo.SysNo, gcnOracle, strErr)
        
        If strErr <> "" Then
            Exit Function
        End If
    End If
    '�걾ID
    'ͼƬ����·��(���������Զ�����),
    '�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
    '��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
    '���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
    '���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
    
    strReturn = mobjImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 0)
    If strReturn = "" Then
        If strErr = "��ͼ�����ݣ�" Then
            strErr = ""
            ReadSampleImage = True
        ElseIf strErr = "" Then
            ReadSampleImage = True
        End If
        Exit Function
    End If
    
    varTmp = Split(strReturn, ",")

    For i = LBound(varTmp) To UBound(varTmp)
        If i > 8 Then Exit For
        If Trim("" & varTmp(i)) <> "" Then
            If Dir(strDir & "\" & Trim("" & varTmp(i))) <> "" Then strChar(i) = strDir & "\" & Trim("" & varTmp(i))
        End If
    Next
    
    ReadSampleImage = True
    Exit Function
errH:
    strErr = "������(ReadSampleImage),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Public Sub FreeImageObj()
    Dim strErr As String
    If Not mobjImg Is Nothing Then
        Call mobjImg.GetSampleImgExit(strErr)
        Set mobjImg = Nothing
    End If
End Sub


