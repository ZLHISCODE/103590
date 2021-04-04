VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cell_Dyn_1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Cell Dyn 1700
'����:��������
'����:�¶�
'�û�:����������Ժ
Private mstrRBC As String, mstrWBC As String, mstrPLT As String '��ͼ�����ݡ�

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
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str���� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim str_Img As String, i As Integer
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, "") - 1)
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            varLine = Split(strLine, ",")
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            If UBound(varLine) >= 73 Then
                If varLine(0) Like "*""   ""*" Then
                    '��������
                    strTmp = Replace(varLine(6), """", "")
                    str���� = Format(CDate("20" & Right(strTmp, 2) & "-" & Replace(Left(strTmp, 5), "/", "-")), "yyyy-MM-dd")
                    strTmp = Replace(varLine(7), """", "")
                    str���� = str���� & " " & strTmp
                    
                    str�걾�� = Val(Replace(varLine(8), """", ""))
                    str�����˱걾�ʿ� = " |ѪҺ|" & IIf(str�걾�� >= 2000, 1, 0)
                    
                    For iRow = 16 To 37 '��16�ŵ�37�� �Ǽ����
                        If Val(varLine(iRow)) <> 0 Then
                            strItem = strItem & "|" & iRow & "|" & Val(varLine(iRow))
                        End If
                    Next
                    
                    If strItem <> "" And Val(str�걾��) <> 0 Then
                        strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                        
                        ''�н������ͼ�����ݸ��ں��档
                        str_Img = ""
                        
                        If mstrWBC <> "" Then
                            If Split(mstrWBC, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '���ڣ��걾����ͬ������ͼ������
                                str_Img = str_Img & "^" & Split(mstrWBC, "|")(1)
                                mstrWBC = ""
                            End If
                        End If
                        
                        If mstrRBC <> "" Then
                            If Split(mstrRBC, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '���ڣ��걾����ͬ������ͼ������
                                str_Img = str_Img & "^" & Split(mstrRBC, "|")(1)
                                mstrRBC = ""
                            End If
                        End If
                        
                        If mstrPLT <> "" Then
                            If Split(mstrPLT, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '���ڣ��걾����ͬ������ͼ������
                                str_Img = str_Img & "^" & Split(mstrPLT, "|")(1)
                                mstrPLT = ""
                            End If
                        End If
                        If str_Img <> "" Then strResult = strResult & vbCrLf & Mid(str_Img, 2)

                    End If
                    
                ElseIf varLine(0) Like "*""PLT""*" Then
                    'PLTͼ�� �ݲ�����
                    If UBound(varLine) >= 272 Then
                        mstrPLT = varLine(6) & varLine(7) & varLine(8) & "|PLT;0;120,262,65-10,130-20,195-30,260-;"
                        For i = 16 To 272
                            mstrPLT = mstrPLT & varLine(i) & ";"
                        Next
                    End If
                ElseIf varLine(0) Like "*""RBC""*" Then
                    
                    'RBCͼ�� �ݲ�����
                    If UBound(varLine) >= 272 Then
                        mstrRBC = varLine(6) & varLine(7) & varLine(8) & "|RBC;0;120,262,52-,104-100,156-,208-200,260-;"
                        For i = 16 To 272
                            mstrRBC = mstrRBC & varLine(i) & ";"
                        Next
                    End If
                ElseIf varLine(0) Like "*""WBC""*" Then
                    'WBCͼ�� �ݲ�����
                    If UBound(varLine) >= 272 Then
                        mstrWBC = varLine(6) & varLine(7) & varLine(8) & "|WBC;0;120,262,43.33-,86.66-100,129.996-,173.326-200,216.656-,259.986-300;"
                        
                        For i = 16 To 272
                            mstrWBC = mstrWBC & varLine(i) & ";"
                        Next
                    End If
                End If
            End If
            If Not strReserved Like "***" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


