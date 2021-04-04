Attribute VB_Name = "mdlDebug"
Option Explicit

Public Const M_STR_MODULE_MENU_TAG As String = "�ɼ�"
Public Const G_STR_HINT_TITLE As String = "��ʾ"
Public Const G_STR_REG_PATH_PUBLIC As String = "����ģ��\zl9PacsCapture"
Public Const G_STR_REG_PATH_PRIVATE As String = "˽��ģ��\zl9PacsCapture\"

Public Enum TDockState                      '��������״̬
    dsClosed = 0    '�ر�
    dsOpen = 1      '��
    dsClosing = 2   '�ر���
End Enum

Public Enum ReportType
    ���Ӳ����༭��
    PACS����༭��
    �����ĵ��༭��
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

Public glngCurVideoContainerHwnd As Long        '��ǰ��Ƶ���ڵ��������ھ��
Public gblnDockingState As TDockState           '�Ƿ��ڵ�������״̬

Public glngInstanceCount As Long
Public gblnOpenDebug As Boolean
Public gobjZOrder As Scripting.Dictionary
Public gblnIsQuitModule As Boolean

Public gstrHotKeyTest As String


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Sub InitCommonLib(cnOracle As ADODB.Connection)
    Set gcnVideoOracle = cnOracle
End Sub

Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    If gblnOpenDebug Then
        OutputDebugString "[" & App.ProductName & "]" & strMethob & "��" & objErr.Description
    End If
End Sub

Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
