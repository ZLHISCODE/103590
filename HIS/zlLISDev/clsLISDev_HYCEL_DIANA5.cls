VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HYCEL_DIANA5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�����Ѫ����
'����:����HYCEL
'����:����
'�û�:�Ͼ���������ҽԺ
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
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '                     a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '                     b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim adata() As String, iDataPoint As Integer, aitem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngloop As Long, intloop As Integer
    Dim intLoopY As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim intStart As Integer             '��ʼλ��
    Dim intEnd As Integer               '����λ��
    Dim strDate As String               '����ʱ��
    
    strCmd = Chr(&H6)
    
    'û���ҵ�������"&H03"ʱ�˳�
    If InStr(strSource, Chr(&H3)) <= 0 Then strReserved = strSource: Exit Sub
    
    
    intStart = InStr(strSource, "") + 1
    intEnd = InStr(strSource, "")
    
    'ֻȡ��һ�ν���������ȷ����´ν���
    strTmp = Mid(strSource, intStart, intEnd - intStart)
    
    strReserved = Mid(strSource, intEnd + 1)
    If InStr(strReserved, Chr(&H2)) <= 0 Then
        'û�п�ʼ��ʱû�з���ֵ
        strReserved = ""
    End If
    
    '��ʼ����
    
    strDate = Mid(strTmp, 4, 4) & "-" & Mid(strTmp, 8, 2) & "-" & Mid(strTmp, 10, 2)
    strData = strDate & "|" & Val(Mid(strTmp, 30, 4)) & "| |Ѫ��|0"
    
    strData = strData & "|1|" & Val(Mid(strTmp, 100, 17)) & "|2|" & Val(Mid(strTmp, 118, 4))
    strData = strData & "|3|" & Val(Mid(strTmp, 122, 5)) & "|4|" & Val(Mid(strTmp, 127, 5))
    strData = strData & "|5|" & Val(Mid(strTmp, 132, 4)) & "|6|" & Val(Mid(strTmp, 136, 4))
    strData = strData & "|7|" & Val(Mid(strTmp, 140, 4)) & "|8|" & Val(Mid(strTmp, 144, 4))
    strData = strData & "|9|" & Val(Mid(strTmp, 148, 4)) & "|10|" & Val(Mid(strTmp, 152, 4))
    strData = strData & "|11|" & Val(Mid(strTmp, 156, 5)) & "|12|" & Val(Mid(strTmp, 161, 5))
    strData = strData & "|13|" & Val(Mid(strTmp, 166, 5)) & "|14|" & Val(Mid(strTmp, 171, 5))
    strData = strData & "|15|" & Val(Mid(strTmp, 176, 5)) & "|16|" & Val(Mid(strTmp, 181, 5))
    strData = strData & "|17|" & Val(Mid(strTmp, 186, 5)) & "|18|" & Val(Mid(strTmp, 191, 4))
    strData = strData & "|19|" & Val(Mid(strTmp, 195, 4)) & "|20|" & Val(Mid(strTmp, 199, 4))
    strData = strData & "|21|" & Val(Mid(strTmp, 203, 4)) & "|22|" & Val(Mid(strTmp, 207, 4))
    
    strResult = strData

End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function
Private Function ReplaceSpace(strData) As String
    Dim intloop As Integer
    Dim strTmp As String
    For intloop = 1 To Len(strData)
        If Mid(strData, intloop, 1) = " " Then
            If Mid(strTmp, Len(strTmp), 1) <> " " Then
                strTmp = strTmp & Mid(strData, intloop)
            End If
        Else
            strTmp = strTmp & Mid(strData, intloop)
        End If
    Next
End Function
