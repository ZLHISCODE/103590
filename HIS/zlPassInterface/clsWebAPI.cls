VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsWebAPI
' Author    : YWJ
' Date      : 2018/11/1
' Purpose   :���ڼ���֪ʶ��WEB����ӿ�
'---------------------------------------------------------------------------------------

Option Explicit

Public Function zlInit(cnMain As ADODB.Connection) As Boolean
'����:
'����:cnMain ���Ӷ���
'    gstrDiagMatchURL = "http://192.168.32.201:8889/bizdomain/b82bd150-09e6-400c-9a60-4fdd3ac9860e"
    Set gcnOracle = cnMain
    
    gstrDiagMatchURL = GetParaURL("֪ʶ��", "���ƥ�����")
    gstrAntidiastoleURL = GetParaURL("֪ʶ��", "������ϲ�ѯ")
    gstrDiagContentURL = GetParaURL("֪ʶ��", "����ĵ���ѯ")
    If gstrDiagMatchURL = "" Or gstrAntidiastoleURL = "" Or gstrDiagContentURL = "" Then Exit Function
    zlInit = True
End Function


Public Function zlShowDiagReference(objfrmMain As Object, ByVal strDiagName As String, Optional ByVal bytStyle As Byte) As Boolean
'����:
'����:strDiagName �������
    Dim strInput As String
    Dim strRet As String
    Dim strMsg As String
    Dim objJSON As Object
    
    On Error GoTo errH
    If gstrDiagMatchURL = "" Or gstrAntidiastoleURL = "" Or gstrDiagContentURL = "" Then Exit Function
    'strInput =[{""name"": ""input_in"",""value"": ""{\""DIA_NAME_QUERY\"": \""�ͼ�Ѫ֢\"",\""NO_QUERY\"": 10}""}]
    
    Set objJSON = mdlJSON.parse("{}")
    Call objJSON.Add("DIA_NAME_QUERY", strDiagName)
    Call objJSON.Add("NO_QUERY", 10)
    strInput = mdlJSON.toString(objJSON)
    Set objJSON = mdlJSON.parse("{}")
    Call objJSON.Add("name", "input_in")
    Call objJSON.Add("value", strInput)
    strInput = "[" & mdlJSON.toString(objJSON) & "]"
    
    strRet = HttpPost(gstrDiagMatchURL, strInput, responseText, , "Basic " & zlStr.Base64Encode("xxx:xxx"))
    '{"output_out":"[{\"DIA_NAME\":\"��������\",\"DIA_ID\":\"d20d3ee3-3c95-4311-9317-be4e8e00c7e7\",\"DIA_NAME_QUERY\":\"����\",\"MATCH_RATE\":0.5},{\"DIA_NAME\":\"��ǻ����\",\"DIA_ID\":\"606e0272-20b4-481a-b8ad-e6f80d19426d\",\"DIA_NAME_QUERY\":\"����\",\"MATCH_RATE\":0.5}]"}
    WriteLog "clsWebAPI", "zlShowDiagReference", "�����ƥ����㡿����ֵ:" & strRet
    If strRet <> "" Then
        Set objJSON = mdlJSON.parse(strRet)
        If objJSON Is Nothing Then
            strMsg = "�����ƥ����㡿����ֵ����ʧ�ܣ�" & mdlJSON.GetParserErrors()
            GoTo errMsg:
        End If
        strRet = NVL(objJSON.Item("output_out"))
        If strRet = "" Then Exit Function
        Set objJSON = parse(strRet)
        If objJSON Is Nothing Then
            strMsg = "�����ƥ����㡿����ֵ����ʧ�ܣ�" & mdlJSON.GetParserErrors()
             GoTo errMsg:
        End If
        frmDiagReference.ShowMe objfrmMain, bytStyle, objJSON
    End If
    zlShowDiagReference = True
    Exit Function
errMsg:
    WriteLog "clsWebAPI", "zlShowDiagReference", strMsg
    Exit Function
errH:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure zlShowDiagReference of Class Module clsWebAPI"
End Function

Private Sub Class_Initialize()
    gsngWaitTime = 3
End Sub