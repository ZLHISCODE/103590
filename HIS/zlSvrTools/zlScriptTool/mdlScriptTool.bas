Attribute VB_Name = "mdlScriptTool"
Option Explicit


'ϵͳ������Ϣ
Public Type SYSPARAM_INFO
    ���ý��С��λ�� As String
    �շ�������Ŀƥ�� As String
    ����Ʊ�ݺų��� As Integer
    �շ�Ʊ�ݺų��� As Integer
    ���￨���볤�� As Integer
    ���￨��ĸǰ׺ As String
    ���￨������ʾ As Boolean
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
    ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
    �շ�Ʊ�� As Integer
    ����Ʊ�� As Integer
    ����Ʊ���ϸ���� As Boolean
    �շ�Ʊ���ϸ���� As Boolean
    ����HIS���� As Byte
End Type

'----------------------------------------------------------------------------------------------------------------------
'ȫ�ֱ�������

Public ParamInfo As SYSPARAM_INFO
Public glngTXTProc As Long                              '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, ParamInfo.ϵͳ��, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************

    On Error GoTo errH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, ParamInfo.ϵͳ��, lngModual, blnSetup)

    Exit Function
    
errH:

End Function


Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.Id
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function
