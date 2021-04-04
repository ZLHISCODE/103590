VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################
Private mclsCommonLib As clsCommonLib

Public Enum ShowConstants '��ʾ����
   WVersteckt = 0    'Versteckt das Fenster
   WNormal = 1       'Zeigt es ganz normal an
   WMaximiert = 3    'Maximiert das Fenster
   WMinimiert = 6    'Minimiert das Fenster
End Enum
Public Enum WaitConstants '�ȴ�����
    wcNone = 0 '���޵ȴ�
    wcInitialisiert = 1
    WCTermination = 2
End Enum
Private Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hwnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type


Private Declare Function WaitForInputIdle Lib "user32" ( _
   ByVal hProcess As Long, _
   ByVal dwMilliseconds As Long _
) As Long '�ȴ��½���������ĳ�ʼ�����ȴ��û����롣

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long _
) As Long '���ȴ����ڹ���״̬ʱ��������رգ���ô������Ϊ��δ����ġ��þ��������� SYNCHRONIZE ����Ȩ�ޡ�

Private Declare Function TerminateProcess Lib "kernel32" ( _
   ByVal hProcess As Long, _
   ByVal uExitCode As Long _
) As Long '��ָֹ�����̼��������̡߳�

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" ( _
   lpExecInfo As SHELLEXECUTEINFO _
) As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function InsertText(ByVal strFile As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT, Optional ByVal dblWaitTime As String = -1) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������strFile:����·��;dblWaitTime:���̵ȴ�ʱ��
    '���أ�
    '******************************************************************************************************************

    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont

    gdblWaitTime = dblWaitTime
    InsertText = AppendPrintData("����", "·��", objFont.HAlignment, objFont.WordWarp, strFile, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)

    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function HavePDF(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ��жϱ������Ƿ�װAdobe Reader�Ķ���
    '������
    '���أ����򷵻�true;���򷵻�false
    '******************************************************************************************************************
    Dim mobjAcroPDF As Object
    Dim i As Integer

'    On Error GoTo errHand
'
'    '����ǰ��ӡ������ΪĬ�ϴ�ӡ��
'    Call SetDefaultPrinter(Printer.DeviceName)
'
'    For i = 0 To frmMain.Controls.Count - 1
'        If frmMain.Controls.Item(i).Name = "oAcroPDF" Then 'me ��������ؼ��Ĵ��������滻��
'            Set mobjAcroPDF = frmMain.Controls.Item(i)
'            HavePDF = True
'            Exit Function
'
'        End If
'    Next
'
'    On Error Resume Next
'    Set mobjAcroPDF = frmMain.Controls.Add("AcroPDF.PDF.1", "oAcroPDF")
'    If Err <> 0 Then
'        HavePDF = False
'        Exit Function
'    End If
    HavePDF = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ShellAndWait(ByVal Operation As String, _
                             ByVal strFilePath As String, _
                             Optional strParameter As String, _
                             Optional strWorkingFolder As String, _
                             Optional WindowSize As ShowConstants = 1, _
                             Optional WaitFor As WaitConstants = 0, _
                             Optional blnCloseProcess As Boolean = False) As String
    '******************************************************************************************************************
    '���ܣ���PDF�ļ�������Ӧ�Ĳ���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngRetVal As Long
    Dim ShExInfo As SHELLEXECUTEINFO

    If strWorkingFolder = "" Then strWorkingFolder = strFilePath
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fMask = &H40
        .hwnd = 0
        .lpVerb = Operation
        .lpFile = strFilePath
        .lpParameters = strParameter
        .lpDirectory = strWorkingFolder
        .nShow = 0
    End With

    Call ShellExecuteEx(ShExInfo)
    
    If WaitFor <> wcNone Then
        gdblWaitTime = 20000
        If WaitFor = wcInitialisiert Then
            lngRetVal = WaitForInputIdle(ShExInfo.hProcess, gdblWaitTime)
        Else
            lngRetVal = WaitForSingleObject(ShExInfo.hProcess, gdblWaitTime)
        End If
        '����,��ȡ������Ϣ
        If lngRetVal = &HFFFFFFFF Then ShellAndWait = Err.Description
    End If

    '�رս���
    If blnCloseProcess = True Then
        lngRetVal = TerminateProcess(ShExInfo.hProcess, 1)
        DoEvents
        If Err.Number <> 0 Then ShellAndWait = "�رս���ʧ�ܣ�" & vbCrLf & "����ԭ��:" & Err.Description
    End If

End Function

Public Sub ShowPDF(ByVal frmMain As Object, ByVal strFileName As String)
    '******************************************************************************************************************
    '���ܣ���ָ��·���µ�PDF�ļ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngBack As Long

    lngBack = ShellExecute(frmMain.hwnd, "open", strFileName, "", "", vbHide)
    If lngBack = 0 Then
        ShowSimpleMsg "�ڴ治��,���ļ�ʧ�ܣ�"
    ElseIf lngBack = 2 Or lngBack = 3 Then
        ShowSimpleMsg "���ܴ��ļ�(""" & strFileName & """),���ļ����ܲ�����!"
    ElseIf lngBack = 26 Then
        ShowSimpleMsg "������������,���ļ�ʧ�ܣ�"
    ElseIf lngBack = 27 Then
        ShowSimpleMsg "���ܴ��ļ�(""" & strFileName & """),���ļ�����ȫ����Ч,���ļ�ʧ��!"
    ElseIf lngBack < 32 Then

        ShowSimpleMsg "���ܴ��ļ�(""" & strFileName & """),�����Ƿ�װAdobe Reader �Ķ���!"
    End If
End Sub

Private Sub Class_Initialize()
    Set mclsCommonLib = New clsCommonLib
End Sub

Private Sub Class_Terminate()
    Set mclsCommonLib = Nothing
End Sub

Public Function FoxitPrint(ByVal strExecute As String) As Boolean
    Dim i As Integer
    Dim strCommand As String
    Dim varFile As Variant
    Dim lngTemp As Long
    
    If strExecute = "" Then Exit Function
    
    Err = 0
    On Error Resume Next
    lngTemp = Shell(strExecute, vbHide)
    Err = 0
    On Error GoTo errHand

    Exit Function

errHand:
End Function

Public Function ShellOpen(ByVal strExecute As String) As Boolean
    Dim strCommand As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    
    Err = 0
    On Error Resume Next
    
    lngTemp = Shell(strExecute, vbHide)
    Err = 0
    On Error GoTo errHand
            
    Exit Function
    
errHand:
    
End Function



