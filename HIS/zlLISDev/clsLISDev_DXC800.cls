VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DXC800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:�ٴ���������ϵ DXC800
'����:��������������� Beckman coulter
'����:�¶�
'�û�:�½�����ľҽԺ

Private mblnStart As Boolean '��ʼ
Private mblnEnd As Boolean   '����
Private mstrData As String '����������

Private mstrLastResult As String '�ϴη��صĴ�
Private mIntLOG As Integer '�ظ�������ͬ���Ĵ���

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
    
    
    '----- �������
    Dim varIn As Variant
    Dim varLin As Variant
    Dim strLine As String, i As Integer
    Dim iStartRow As Integer, iEndRow As Integer, str���� As String, strֵ As String
    Dim StrInput As String, strָ�� As String
    Dim curTP As Currency, curALB As Currency, curTBIL As Currency, curDBIL As Currency
    
    On Error GoTo errHandle
    
    StrInput = Replace(strSource, vbCrLf, "��" & vbCrLf)
    '----  ��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    iStartRow = -1: iEndRow = -1
    
    If StrInput Like "*[[]*[]]*" Or StrInput Like "[[]*[]]*" Then
        varIn = Split(StrInput, vbCrLf)
        
        For i = LBound(varIn) To UBound(varIn)
            strLine = varIn(i)
            If strLine Like "*[[]*,802,03,*[]]*" Or strLine Like "[[]*,802,03,*[]]*" Then
                
                strLine = Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1)
                varLin = Split(strLine, ",")
                str���� = Mid(varLin(3), 5, 4) & "-" & Mid(varLin(3), 3, 2) & "-" & Mid(varLin(3), 1, 2)
                
                strֵ = IIf(Trim(varLin(15)) Like "*[#]*", " ", Trim(varLin(15)))
                strָ�� = Trim(varLin(10))
                If strָ�� = "" Then strָ�� = " "
                If strָ�� = "07D" Then curTP = Val(strֵ)
                If strָ�� = "08D" Then curALB = Val(strֵ)
                If strָ�� = "11A" Then curTBIL = Val(strֵ)
                If strָ�� = "11B" Then curDBIL = Val(strֵ)
                strResult = strResult & str���� & "|" & Val(varLin(9)) & "| |ѪҺ|0|" & strָ�� & "|" & IIf(strֵ = "", " ", strֵ) & "||"
                
                If curTP <> 0 And curALB <> 0 Then
                    strResult = strResult & str���� & "|" & Val(varLin(9)) & "| |ѪҺ|0|GLO|" & Format((curTP - curALB), "0.00") & "|A/G|" & Format(curALB / (curTP - curALB), "0.00") & "||"
                    curTP = 0
                    curALB = 0
                End If
                
                If curTBIL <> 0 And curDBIL <> 0 Then
                    strResult = strResult & str���� & "|" & Val(varLin(9)) & "| |ѪҺ|0|IBIL|" & Format((curTBIL - curDBIL), "0.00") & "||"
                    curTBIL = 0
                    curDBIL = 0
                End If
                
            ElseIf strLine Like "[[]*,802,01,*[]]*" Or strLine Like "*[[]*,802,01,*[]]*" Then
                'һ���걾��ʼ
                iStartRow = i
            ElseIf strLine Like "*[[]*,802,05,*[]]*" Or strLine Like "[[]*,802,05,*[]]*" Then
                '��¼�ѽ�����
                iEndRow = i
            Else
                If Not strLine Like "*[[]*[]]*" Or strLine Like "[[]*[]]*" Then
                    '���ز�������������
                    If strLine <> "" Then
                        strReserved = strReserved & Replace(strLine, "��", vbCrLf)
                    End If
                End If
            End If
        Next
        
    Else
        '���ݲ�����,���ش���ֵ
        strReserved = strSource
    End If
    
    If Len(strResult) > 0 Then
        If Mid(strResult, 1, 2) = "||" Then strResult = Mid(strResult, 3)
        If Right(strResult, 2) = "||" Then strResult = Mid(strResult, 1, Len(strResult) - 2)
    End If
    
    If strResult <> "" Then mIntLOG = 0 '�н������,�Ͱ��ظ�������ͬ���Ĵ�������
    
    If mstrLastResult <> "" And strReserved <> "" Then
        If mstrLastResult = strReserved And mstrLastResult = strSource And strResult = "" Then
            mIntLOG = mIntLOG + 1
        End If
    End If
    mstrLastResult = strReserved
    
    '--- �ظ�������100����ͬ�Ĵ�,û��û�ý��,��������
'    If mIntLOG > 10000 Then
'        strReserved = ""
'        mIntLOG = 0
'    End If
    
    If mIntLOG < 5 Then
        '���ظ�д��־
        Call WriteLog("DXC800 analyse", "strsource:" & strSource & vbNewLine & "strReserved:" & strReserved, "strResult:" & strResult)
    End If
    Exit Sub
errHandle:
    Call WriteLog("DXC800 analyse", "�������" & Err.Description, "")
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

