Attribute VB_Name = "mdlQueue"
Option Explicit

Public Const MS_SOUND_TYPE = "΢������"
Public Const DEF_SOUND_TYPE = "ϵͳĬ��"
Private Const ISDEBUG = False 'True
Public Const G_LNG_QUEUEMANAGE_MODULENUM = 1160

'�����п��ʼ��
Public Const C_STR_QUEUECALL = "0,0,0,0,50,0,90,0,60,0,0,60,60,0,0,60,0,0,125"
'�Ŷ��п��ʼ��
Public Const C_STR_QUEUEQUEUE = "0,0,0,30,50,0,90,40,60,60,0,60,60,50,125,0,120,60,0"

'LED��ʾ��ر���
Public plngLEDModal As Long                'LEDģ�����
Public prsLEDComponent As New ADODB.Recordset  'LED���������ݿ���Ϣ
Public pobjLEDShow As Object               'LED����
Public glngSys As Long
Public glngModul As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'��������ַ���
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'�������ŵĺ���
Public Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Public Declare Function StopPlayStr Lib "StrSound" () As Long
Public Declare Function InitHtTextSound Lib "StrSound.dll" () As Boolean
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


'�ж������Ƿ�Ϊ��
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long


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

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
errHand:
    Substr = ""
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function
