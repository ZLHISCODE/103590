Attribute VB_Name = "mdlDebug"
Option Explicit

#Const DebugVer = True

Public Const M_STR_MODULE_MENU_TAG As String = "�ɼ�"
Public Const G_STR_HINT_TITLE As String = "��ʾ"

Public Const G_STR_REG_PATH_PUBLIC As String = "����ģ��\zl9PacsCapture"
Public Const G_STR_REG_PATH_PRIVATE As String = "˽��ģ��\zl9PacsCapture\"


Public Enum TDockState                      '��������״̬
    dsClosed = 0    '�ر�
    dsOpen = 1      '��
    dsClosing = 2   '�ر���
End Enum


Public gcnVideoOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��

Public gobjOwner As Object
Public glngRootHandle As Long
Public glngSys As Long
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModule As Long                   'ģ���
Public glngDepartId As Long                 '��ǰ����ID

Public gobjCapturePar As clsCaptureParameter    '��Ƶ�ɼ���ز�������

Public gobjNotifyEvent As clsNotifyEvent         '��Ϣ֪ͨ��������
Public gobjVideo As frmWork_Video                '��Ƶ�ɼ�����
'Public gobjGlobal As clsGlobal

Public glngCurVideoContainerHwnd As Long        '��ǰ��Ƶ���ڵ��������ھ��
'Public glngNextVideoContainerHwnd As Long       '��ǰ��Ƶ���ڴ���Z���е���һ���д��ھ��
Public gblnDockingState As TDockState           '�Ƿ��ڵ�������״̬

Public glngInstanceCount As Long
Public gblnOpenDebug As Boolean
Public gobjZOrder As Scripting.Dictionary
Public gblnIsQuitModule As Boolean

'debug property
Public gstrHotKeyTest As String




Private Function IsDebugMode() As Boolean

    IsDebugMode = False

    On Error Resume Next

    Debug.Print 1 / 0

    If err.Number <> 0 Then

        IsDebugMode = True

    End If

End Function



Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Sub Main()
BUGEX "Main 1", True
    If UCase(Command()) = "DEBUG" Or IsDebugMode Then
BUGEX "Main Enter Debug", True
        frmTestLogin.Show
    End If

BUGEX "Main End", True




'    If Not IsDebugMode Then Exit Sub
'
'    Set gcnOracle = New ADODB.Connection
'
'    OraDataOpen "", "zlhis", "HIS"
'
'    Set gcnVideoOracle = gcnOracle
'
'    frmTestWindow.Show

End Sub


Public Sub InitCommonLib(cnOracle As ADODB.Connection)
'��ʼ���������(���ڽ�������Ŀ)
    Dim blnIsEqualDB As Boolean
    
    If cnOracle Is Nothing Then Exit Sub
    
 
    If gobjComLib Is Nothing Then
        'Set gobjComLib = zl9ComLib.clsComLib
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    End If

    blnIsEqualDB = False
    If Not gcnVideoOracle Is Nothing Then
        blnIsEqualDB = IIf(gcnVideoOracle.ConnectionString = cnOracle.ConnectionString, True, False)
    End If
    
    '������Ӳ�ͬ������Ҫ���´�������
    If Not blnIsEqualDB Then
        Set gcnVideoOracle = Nothing
        
        '�����ݿ����Ӹı�ʱ�����´�������
        Set gcnVideoOracle = New ADODB.Connection
            
        'ע����������ActiveExeΪ�����Ľ�����Ŀ����˲���ʹ��cnOracleֱ�Ӷ�gcnVideoOracle����ֵ������������������Ͳ���ȷ,XXX���Ĵ���
        gcnVideoOracle.ConnectionString = cnOracle.ConnectionString
            
        '�����ݿ�����
        gcnVideoOracle.Open
    Else
        Exit Sub
    End If
    
    If gobjComLib.gstrNodeNo <> "" Then Exit Sub
    
    Call zlCL_InitCommon(gcnVideoOracle)
    Call zlCl_RegCheck
End Sub


Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    If gblnOpenDebug Then
        OutputDebugString "[" & App.ProductName & "]" & strMethob & "��" & objErr.Description
    End If
End Sub


'Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
'    If gblnOpenDebug Or blnIsForce Then
'        OutputDebugString Now & " |---> " & strDebug
'    End If
'End Sub

Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
