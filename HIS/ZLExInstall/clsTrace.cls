VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'**ģ �� ����clsTrace
'**�� �� �ˣ�lshuo
'**��    �ڣ�2016/10/11
'**��    ��������ģ���ṩ����־�����࣬����֧����־����
'**��    ����V1.0
'*************************************************************************
'===========================================================================
'                0��ģ������빫������
'===========================================================================
Public Enum LogTimeType                                 '��־ʱ������
    LTT_None = 0                                        '������ʱ��
    LTT_FullDate = 1                                    'ȫ����ʱ���ʽ
    LTT_OnlyTime = 2                                    'ֻ��ʱ��
End Enum
Public Enum SectionLevel                                '�μ���
    SL_LevelOne = 0                                     'һ����
    SL_LevelTwo = 1                                     '������
    SL_LevelThree = 2                                   '������
End Enum
Private mobjLog                 As TextStream
Private Const INFO_LENGTH       As Long = 15
Private mlngLoop                As Long
'===========================================================================
'                1���������Եı���
'===========================================================================
'LogFile��־�ļ��ļ�����(R)
Private mstrLogFile             As String
'===========================================================================
'                2�����Զ���
'===========================================================================
'LogFile��־�ļ��ļ�����(R)
Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'===========================================================================
'                3����������
'===========================================================================
Public Function OpenTace(ByVal strTaceName As String, Optional ByVal strSetupPath As String) As Boolean
    Dim objFSO      As New FileSystemObject, strFolder      As String
    
    On Error GoTo ErrH
    If Not objFSO.FolderExists(strSetupPath) Then
        Call objFSO.CreateFolder(strSetupPath)
    End If
    strFolder = strSetupPath & "\Log"
    If Not objFSO.FolderExists(strFolder) Then
        Call objFSO.CreateFolder(strFolder)
    End If
    strFolder = strFolder & "\��־����"
    If Not objFSO.FolderExists(strFolder) Then
        Call objFSO.CreateFolder(strFolder)
    End If
    mstrLogFile = strFolder & "\" & strTaceName & "_" & Format(Now, "YYMMDDHHmmss") & ".log"
    Set mobjLog = objFSO.CreateTextFile(mstrLogFile, True)
    OpenTace = True
    Exit Function
ErrH:
    mstrLogFile = ""
    MsgBox "����־����ʧ�ܣ�������Ϣ��" & vbNewLine & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Public Sub WriteSection(ByVal strSectionName As String, Optional ByVal slLevel As SectionLevel = SL_LevelOne)
'���ܣ���־��д��һ���׶�
'������strSectionName=����
    If Not mobjLog Is Nothing Then
        If slLevel = SL_LevelOne Then
            mobjLog.WriteLine String(120, "*")
            mobjLog.WriteLine String(4, "*") & LPAD(strSectionName, (120 - 4 * 2 + Len(strSectionName)) \ 2)
            mobjLog.WriteLine String(120, "*")
        ElseIf slLevel = SL_LevelTwo Then
            mobjLog.WriteLine String(100, "=")
            mobjLog.WriteLine String(4, "=") & LPAD(strSectionName, (100 - 4 * 2 + Len(strSectionName)) \ 2)
            mobjLog.WriteLine String(100, "=")
        ElseIf slLevel = SL_LevelThree Then
            mobjLog.WriteLine RPAD(LPAD(strSectionName, (50 + ActualLen(strSectionName) \ 2), "-"), 100, "-")
        End If
    End If
End Sub

Public Sub WriteInfo(ByVal strInfoName As String, ParamArray arrInfo() As Variant)
'���ܣ���־��һ����Ϣ������Ϣ��Ϊ�������
    mobjLog.Write "[" & RPAD(strInfoName, INFO_LENGTH) & "]"
    For mlngLoop = LBound(arrInfo) To UBound(arrInfo) Step 2
        mobjLog.Write arrInfo(mlngLoop) & ":" & arrInfo(mlngLoop + 1) & "    "
    Next
    mobjLog.WriteLine
End Sub

'��־����
Public Sub WriteLog(Optional ByVal strText As String, Optional ByVal lttAddTime As LogTimeType = LTT_None, Optional ByVal lngReturnLines As Long)
'����:strText       =Ҫд���һ����־�ı�,���Ϊ�ձ�ʾдһ�л��з�
'     bytAddTime    >0ʱ������־�ı�֮ǰ������־ʱ�䣬1=���ں�ʱ��������ʽ,2-��ʱ��,0-������
'     lngReturnLines=����־�ı�֮��д�����л��з�,0-��д���з�
    If Not mobjLog Is Nothing Then
        '������־ʱ��
        If lttAddTime <> LTT_None Then strText = LogTime(lttAddTime = LTT_OnlyTime) & strText
        '��ȡ��������
        If Len(strText) > 500 Then strText = Mid(strText, 1, 500)
        'д��־�ı�
        If strText = "" Then
            mobjLog.WriteBlankLines 1
        Else
            mobjLog.WriteLine strText
            If lngReturnLines > 0 Then mobjLog.WriteBlankLines lngReturnLines
        End If
    End If
End Sub

Public Sub CloseLog()
    '�ر���־�ļ�
    mobjLog.Close
    Set mobjLog = Nothing
End Sub
'===========================================================================
'                4��˽�з���
'===========================================================================
Private Function LogTime(Optional blnOnlyTime As Boolean) As String
    If blnOnlyTime Then
        LogTime = Format(Now, "HH:mm:ss")
    Else
        LogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
    End If
End Function

Private Function LPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = ActualLen(strCode)
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = strCode
    End If
    LPAD = Replace(strTmp, Chr(0), strChar)
End Function

Private Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = ActualLen(strCode)
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = strCode
    End If
    'ȡ��������ַ�
    RPAD = Replace(strTmp, Chr(0), strChar)
End Function
