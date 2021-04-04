Attribute VB_Name = "mdlQueue"
Option Explicit

Private Const ISDEBUG = False 'True

'��Ϣ��Ƕ���
Public Const G_STR_MSG_QUEUE_001 As String = "ZLHIS_QUEUE_001" '�����Ϣ
Public Const G_STR_MSG_QUEUE_002 As String = "ZLHIS_QUEUE_002" '�����Ϣ
Public Const G_STR_MSG_QUEUE_003 As String = "ZLHIS_QUEUE_003" '״̬ͬ��
Public Const G_STR_MSG_QUEUE_004 As String = "ZLHIS_QUEUE_004" '��������

'����߱����ж���
Public Const G_STR_MUST_NEED_QUEUE_COL As String = "ID,��������,ҵ��ID,��������,�ŶӺ���,�Ŷ�״̬,�Ŷ����"

Public gstrRegPath As String

Public gobjMsgCenter As clsQueueMsgCenter


Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
   


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'��������ַ���
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)


'�������ŵĺ���
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


'�ж������Ƿ�Ϊ��
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long


'����GUID
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As TGUID) As Long
   
   
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
   
   
Public Function GetGUID() As String
'��ȡGUID��
    Dim udtGuid As TGUID
    
    If (CoCreateGuid(udtGuid) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGuid.Data1)), "0") & Hex$(udtGuid.Data1) & _
        String(4 - Len(Hex$(udtGuid.Data2)), "0") & Hex$(udtGuid.Data2) & _
        String(4 - Len(Hex$(udtGuid.Data3)), "0") & Hex$(udtGuid.Data3) & _
        IIf((udtGuid.Data4(0) < &H10), "0", "") & Hex$(udtGuid.Data4(0)) & _
        IIf((udtGuid.Data4(1) < &H10), "0", "") & Hex$(udtGuid.Data4(1)) & _
        IIf((udtGuid.Data4(2) < &H10), "0", "") & Hex$(udtGuid.Data4(2)) & _
        IIf((udtGuid.Data4(3) < &H10), "0", "") & Hex$(udtGuid.Data4(3)) & _
        IIf((udtGuid.Data4(4) < &H10), "0", "") & Hex$(udtGuid.Data4(4)) & _
        IIf((udtGuid.Data4(5) < &H10), "0", "") & Hex$(udtGuid.Data4(5)) & _
        IIf((udtGuid.Data4(6) < &H10), "0", "") & Hex$(udtGuid.Data4(6)) & _
        IIf((udtGuid.Data4(7) < &H10), "0", "") & Hex$(udtGuid.Data4(7))
    End If
End Function
Public Sub DebugMsg(ByVal strClassName, strMethod, strParameter As String, Optional strExceptionMsg As String = "")
    If ISDEBUG Then
        Call OutputDebugString(Now & ">> [���ù��̣�" & strClassName & "." & strMethod & "]  [�������ݣ�" & strParameter & "]  " & _
                                IIf(Trim(strExceptionMsg) <> "", "[�쳣��Ϣ��" & strExceptionMsg & "]", ""))
    End If
End Sub


Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function StrNvl(ByVal varValue As String, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    StrNvl = IIf(Trim(varValue) = "", DefaultValue, varValue)
End Function

Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function



Public Function GetColIndex(strColumnName As String, objQueueList As Object) As Long
'���ݴ���������õ�����Index
On Error GoTo errHandle
    Dim i As Integer
    Dim objCurFindQueueList As ReportControl
    
    GetColIndex = -1
    
    If objQueueList Is Nothing Then Exit Function
    
    Set objCurFindQueueList = objQueueList
    
    With objCurFindQueueList
    
        For i = 0 To .Columns.Count - 1
            If .Columns(i).Caption = strColumnName Then
                GetColIndex = .Columns(i).ItemIndex
                Exit Function
            End If
        Next i
    
    End With
    
Exit Function
errHandle:
    GetColIndex = -1
End Function


Public Function HasField(rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean
'�ж�ado���Ƿ����ָ���ֶ�
    Dim i As Long
    
    HasField = False
    
    For i = 0 To rsData.Fields.Count - 1
        If UCase(rsData.Fields(i).Name) = strFieldName Then
            HasField = True
            Exit Function
        End If
    Next i
End Function
