Attribute VB_Name = "mdlDocShiftBase"
Option Explicit
Public Const conMenu_DocShift_FilePopup = 1              '�ļ�
Public Const conMenu_DocShift_PatiTypePopup = 2            '��������
Public Const conMenu_DocShift_PatiProjectPopup = 3              '������Ŀ
Public Const conMenu_DocShift_ViewPopup = 7             '�鿴
Public Const conMenu_DocShift_HelpPopup = 9            '����

'�ļ��˵�
Public Const conMenu_DocShift_File_Preview = 101         'Ԥ��(&V)
Public Const conMenu_DocShift_File_Exit = 191            '�˳�(&X)

'�������Ͳ˵�
Public Const conMenu_DocShift_Edit_New = 201         '�²�������(&A)
Public Const conMenu_DocShift_Edit_Modify = 202          '�޸�(&M)
Public Const conMenu_DocShift_Edit_Delete = 203          'ɾ��(&D)
Public Const conMenu_DocShift_Edit_Reuse = 204          '����(&D)
Public Const conMenu_DocShift_Edit_Stop = 205          'ͣ��(&D)

'������Ŀ�˵�
Public Const conMenu_DocShift_Edit_NewProject = 301        '�²�����Ŀ(&A)
Public Const conMenu_DocShift_Edit_ModifyProject = 302          '�޸�(&M)
Public Const conMenu_DocShift_Edit_DeleteProject = 303          'ɾ��(&D)
Public Const conMenu_DocShift_Edit_RowProject = 304        '�����ͬ����Ŀ��ϲ�

'�鿴�˵�
Public Const conMenu_DocShift_View_ToolBar = 701              '������(&T)
Public Const conMenu_DocShift_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_DocShift_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_DocShift_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_DocShift_View_StatusBar = 702            '״̬��(&S)

'�����˵�
Public Const conMenu_DocShift_Help_Help = 901        '��������(&H)
Public Const conMenu_DocShift_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_DocShift_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_DocShift_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_DocShift_Help_Web_Mail = 9022       '���ͷ���(&M)
Public Const conMenu_DocShift_Help_About = 991       '����(&A)��

Public Function rsPatiType(ByVal strSName As String) As ADODB.Recordset
'���ݲ��˼�ƻ�ȡ����������Ϣ
    
    On Error GoTo errH
    gstrSql = "Select ���, ����, ˳��,��ʼ����, ��ȡsql, �Ƿ�ͣ�� From ҽ�����Ӱಡ������ Where ��� = [1]"
    Set rsPatiType = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����������Ϣ", strSName)
    Exit Function
errH:
    MsgBox err.Description, vbInformation, "��ȡ����������Ϣ"
End Function

Public Function GetPatiTypeInfo(ByVal strType As String, Optional strPatiTypeInfo As String) As ADODB.Recordset
'���ݲ������ͼ�ƻ�ȡ����������Ϣ
    
    gstrSql = ""
    If strPatiTypeInfo <> "" Then gstrSql = " And ��Ŀ����=[2]"
    gstrSql = "Select ���˼��, ��Ŀ����, ���, ��Ŀ���, Decode(������ʽ, 1, '1-�����', 2, '2-����ѡ��', 3, '3-����ѡ��') ������ʽ," & vbNewLine & _
        "       Decode(Nvl(��������,0), 0, '0-�ı�', 1, '1-����', 2, '2-����') ��������, �����ʽ, ����ֵ��, ��������," & vbNewLine & _
        "       Decode(��ȡ��Դ, 1, '1-�������', 2, '2-��������', 3, '3-��Ѫ���', 4, '4-��������', '99', '99-SQL��ȡ') ��ȡ��Դ, ��ȡ����, ��ȡsql, ��������, �Ƿ�ֻ��, ����������" & vbNewLine & _
        "From ҽ�����Ӱಡ����Ŀ" & vbNewLine & _
        "Where ���˼�� = [1]" & gstrSql & vbNewLine & _
        "Order By ���"
    Set GetPatiTypeInfo = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����������Ϣ", strType, strPatiTypeInfo)
End Function

Public Function GetSqlColor() As String
    Dim objfso As New FileSystemObject
    '��������:��ȡ�﷨�ؼ���SQL�﷨������ʾ����
    '��ȡ��ֱ�ӽ��﷨�ؼ���SyntaxScheme������Ϊ����ֵ����
    Dim strColor As String, strPath As String
    
    strPath = objfso.GetParentFolderName(GetSetting("ZLSOFT", "����ȫ��", "����·��")) & "\PUBLIC\_sql.schclass"
    If Not objfso.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    
    If objfso.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    End If
    GetSqlColor = strColor
End Function

Public Function ReadFileToString(ByVal strFile As String) As String
    Dim strBuffer As String
    Dim lngHwnd As Long
    Dim lngFileLen As Long

    lngHwnd = FreeFile

    On Error Resume Next
    Open strFile For Binary Shared As lngHwnd
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & vbCrLf & err.Description & vbCrLf & "Error in ReadFileToString, File='" & strFile & "'", vbCritical
        GoTo Proc_Exit
    End If
    On Error GoTo 0
    
    lngFileLen = LOF(lngHwnd)
    strBuffer = Space(lngFileLen)
    Get lngHwnd, , strBuffer
    
    Close lngHwnd
    
Proc_Exit:
    ReadFileToString = strBuffer
End Function


