VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_FUJI_FDC3500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FUJI_FDC3500
'����:FUJI
'����:����
'�û�:��ͩ
Private mblnStart As Boolean '��ʼ
Private mblnEnd As Boolean   '����
Private mstrData As String '����������

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
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '�����������
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String, strֵ As String
    Dim strҪ��ȡָ�� As String, varҪ��ȡָ�� As Variant
    Dim iStrPoint As Integer
    Dim strNow As String
    Dim aItem() As String
    Dim strData As String
    Dim intLoop As Integer
    
    If InStr(strSource, "") = 0 Or InStr(strSource, "") = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, "") + 1, InStr(strSource, "") - 1)
    
    strNow = DelSapce(strNow)
    
    aItem = Split(strNow, " ")
    
    strData = Mid(aItem(1), 1, 10) & " " & Mid(aItem(1), 11, 5) & ":00|" & Mid(aItem(1), 16, 1) & "| |Ѫ��|0"
    
    For intLoop = 0 To UBound(aItem)
        If InStr(aItem(intLoop), "=") > 0 Then
            strData = strData & "|" & aItem(intLoop - 1) & "|" & Val(Replace(aItem(intLoop), "=", ""))
        End If
    Next
    
    strResult = strData
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

Private Function DelSapce(strLine As String) As String
    '����       ɾ������Ŀո�
    Dim intLoop  As Integer
    Dim strNow As String
    strNow = strLine
    For intLoop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intLoop), Space(1))
    Next
    DelSapce = strNow
End Function
'------------------------------------------------------------------------------------------------------------
'    Dim iStrPoint As Integer '�ַ�ָ��
'    Dim strBuffer As String, strData As String
'    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
'
'    Dim varIn As Variant '�����������
'    Dim strTime As String
'
'    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
'    Dim strLine As String
'    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String, strֵ As String
'    Dim strҪ��ȡָ�� As String, varҪ��ȡָ�� As Variant
'
'    On Error GoTo errHandle
'    str�����˱걾�ʿ� = " |ѪҺ|0"
'    '��ʼ������ֵ
'    strReserved = "": strResult = "": strCmd = ""
'    '-- ����Ҫ��ȡ��ָ��
'    strҪ��ȡָ�� = "Temp;FIO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
'    varҪ��ȡָ�� = Split(strҪ��ȡָ��, ";")
'
'    If strSource = "" Then Exit Sub
'    Call WriteLog("Abl5.Analyse ���� ", strSource, mstrData)
'    varIn = Split(strSource, vbNewLine)
'    If UBound(varIn) <= 0 Then
'        varIn = Split(strSource, Chr(10))
'    End If
'
'    If mblnStart = True And mblnEnd = False Then
'            'strReserved = strSource
'        mstrData = mstrData & strSource
'        Call WriteLog("Abl5.Analyse ��ȡ���� ", strSource, mstrData)
'        If mstrData Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse �ҵ������� ", strSource, mstrData)
'        End If
'    End If
'
'
'    If strSource Like "*ABL5*" Then
'        mblnStart = True
'        'strReserved = strSource
'        Call WriteLog("Abl5.Analyse �ҵ���ʼ�� ", strSource, mstrData)
'        If strSource Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse ȫ������,�ҵ������� ", strSource, mstrData)
'        End If
'        mstrData = strSource
'    End If
'
'    If mblnStart And mblnEnd And mstrData <> "" Then
'        Call WriteLog("Abl5.Analyse ��ʼ���� ", strSource, mstrData)
'    End If
'    Exit Sub
'errHandle:
'    Call WriteLog("Abl5.Analyse �������", strSource, strResult & vbNewLine & Err.Description)

