VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_AU680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'ɽ������������
'au680 ������  ����
'2011-11-21 �¶�
'�޸ģ�������
'������ɽ������ҽԺ�����ǺϾ�����Ѫ�ܲ�ҽԺ,��ˮ�ڶ�����ҽԺ
'2012-12-26

Private mstrCode As String
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim lng_S As Long, lng_E As Long
    Dim strCode As String, strLine As String, strTmp As String, strItem As String
    Dim lngSampleS As Long, lngSampleLen As Long
    Dim lngItemS As Long, lngItemLen As Long, lngItemNoLen As Long
    Dim strSampeNo As String, strDate As String
    Dim str_�������� As String
    Dim intE As Integer, intQ As Integer, intP As Integer
    lngSampleS = 11     '�걾�ſ�ʼλ��
    lngSampleLen = 4    '�걾�ų���
    lngItemS = 20       '��Ŀ��ʼλ��
    lngItemLen = 11     'һ����Ŀ����ĳ���
    lngItemNoLen = 3    'ͨ���볤��
    
    
        
    intE = Val(ReadIni("AU680", "E", App.Path & "\AU680.ini"))
    If intE < 0 Then intE = 8000:   Call WriteIni("AU680", "E", 8000, App.Path & "\AU680.ini")
    
    intQ = Val(ReadIni("AU680", "Q", App.Path & "\AU680.ini"))
    If intQ < 0 Then intQ = 6000:   Call WriteIni("AU680", "Q", 6000, App.Path & "\AU680.ini")
    
    intP = Val(ReadIni("AU680", "P", App.Path & "\AU680.ini"))
    If intP < 0 Then intP = 3000:   Call WriteIni("AU680", "P", 3000, App.Path & "\AU680.ini")
    str_�������� = Trim(ReadIni("��������", "����", App.Path & "\AU680.ini"))
    If str_�������� = "" Then
        WriteIni "��������", "����", "AU680", App.Path & "\AU680.ini"
        str_�������� = "AU680"
    End If
    
    If str_�������� = "AU680_LSXRMYY" Then
        lngSampleS = 13
        lngItemS = 48
        lngItemLen = 14
    ElseIf str_�������� = "AU680_TSRMYY" Then
        lngItemS = 46
        lngItemLen = 14
    End If

    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    mstrCode = mstrCode & strSource
    
    strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Do While strCode <> ""
'        If strCode Like Chr(&H2) & "DE*" Then
            
        strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strSampeNo = Mid(strCode, lngSampleS, lngSampleLen)
        If Not IsNumeric(Mid(strSampeNo, 1, 1)) Then
            Select Case Mid(strSampeNo, 1, 1)
                Case "P"
                    strSampeNo = intP + Mid(strSampeNo, 2)
                Case "E"
                    strSampeNo = intE + Mid(strSampeNo, 2)
            End Select
        End If
        strLine = Mid$(strCode, lngItemS)
        strItem = ""
        Do While Len(strLine) > lngItemLen
            
            strTmp = Mid(strLine, 1, lngItemLen)
            strLine = Mid(strLine, lngItemLen + 1)
            If Trim(Mid(strTmp, 1, lngItemNoLen)) <> "" And Trim(Mid(strTmp, lngItemNoLen + 1)) <> "" Then
                strItem = strItem & "|" & Mid(strTmp, 1, lngItemNoLen) & "|" & Val(Mid$(strTmp, lngItemNoLen + 1))
            End If
            
        Loop
        If Val(strSampeNo) > 0 And strItem <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampeNo & "| |Ѫ��|0" & strItem
        End If
'        End If
        strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

End Function












