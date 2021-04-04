VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC5500_bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BC-5500
'����:��������
'����:�¶�
'�û�:�����һ����ҽԺ

Private mstrIn As String
Const strRetuItems = ",WBC,Neu#,Lymph#,Mon#,Eos#,Bas#,Neu%,Lymph%,Mon%,Eos%,Bas%,RBC,HGB,HCT,MCV,MCH,MCHC,RDW-CV,RDW-SD,PLT,MPV,PDW,PCT,"  'Ҫ��������Ŀ

Dim StrInput As String

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
    
    Dim strData As String, lngS As Long, lngE As Long
    Dim str�걾�� As String, str���� As String, str��Ŀ As String
    Dim str�ʿ� As String, strItemName As String

    Dim strSegment As String '��һ�� Segment������
    Dim lngSD_FD As Long '��SD��FD�ķָ���λ��
    Dim lngFD_V As Long '�� FD��V�ķָ���λ��
    Dim strFD_Name As String '��FD��˵��
    
    Dim strSegName As String '��segment��˵��
    
    Dim lngImgDataLen As Long, lngImgStart As Long, lngI As Long
    Dim strImgRBC As String, strImgPLT As String, strImgDiff As String, strImgBASO As String
    Dim strImgData As String, strTmp�걾�� As String, i As Integer
    If strSource = ",02" Then
        strCmd = ",06" ' Chr(&H6)
    Else
        StrInput = StrInput & strSource
    End If
    
    If StrInput Like "*,05,##,##,##,03,*,08,04,0A*" Then
        '---------
        str�걾�� = "": str���� = "": str��Ŀ = "": strResult = "": strData = ""
        strImgRBC = "": strImgPLT = "": strImgDiff = "": strImgBASO = ""
        lngImgDataLen = 0: lngImgStart = 0: lngI = 0
        '-----
        lngS = Instr_Bin(StrInput, Chr(&H5)) '--��ʼ�� MS
        If Mid_bin(StrInput, lngS + 4, 1) <> Chr(&H3) Then
            '���ǿ�ʼλ�ã��˳�
            StrInput = Mid_bin(StrInput, lngS + 1, , False)
            Exit Sub
        End If
        
        lngE = Instr_Bin(StrInput, Chr(&H8) & Chr(&H4) & Chr(&HA))        '--������         chr(10)
        
        If lngE <= lngS Then Exit Sub
        
        strData = Mid_bin(StrInput, lngS, lngE - lngS + 3, False)
        StrInput = Mid_bin(StrInput, lngE + 3, , False)
        'If Instr_Bin(strInput, Chr(&H5)) <= 0 Then strReserved = ""
        
        lngS = 1
        
        str�ʿ� = UCase(Mid_bin(strData, lngS + 1, 3))
        If str�ʿ� = "CTR" Then
            str�ʿ� = "0"
            strData = Mid_bin(strData, lngS + 5, , False)
        ElseIf str�ʿ� = "QCR" Then
            str�ʿ� = "1"
            strData = Mid_bin(strData, lngS + 5, , False)
        ElseIf str�ʿ� = "QCC" Then
            '������ݲ�����
            Exit Sub
        End If
        
        Do While Len(strData) > 0
            '��һ������
            lngE = Instr_Bin(strData, Chr(&H4))
            If lngE > 0 Then
                strSegment = Mid_bin(strData, 1, lngE - 1, False)
            Else
                '���һ��
                strSegment = strData
            End If
            
            lngSD_FD = Instr_Bin(strSegment, Chr(&HC))
            
            If lngSD_FD > 0 Then
                strSegName = Mid_bin(strSegment, 1, lngSD_FD - 1)
                strSegment = Mid_bin(strSegment, lngSD_FD + 1, , False)
                If InStr(strRetuItems, "," & strSegName & ",") > 0 Then     '������Ŀ
                    
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '�ı����ݣ�����ֱ����strsegment������
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '��ʽ����,������
                            strSegment = ""
                        End If
                        
                        If strFD_Name = "Val" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str��Ŀ = str��Ŀ & "|" & strSegName & "|" & Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str��Ŀ = str��Ŀ & "|" & strSegName & "|" & Mid_bin(strSegment, 1)
                            End If
                            strSegment = ""
                            
                        End If
                    Loop
                ElseIf strSegName = "PatInfo" Then                          '�걾��
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '�ı����ݣ�����ֱ����strsegment������
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '��ʽ����,������
                            strSegment = ""
                        End If

                        
                        If strFD_Name = "SampleID" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str�걾�� = Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str�걾�� = Val(Mid_bin(strSegment, 1))
                            End If
                            If Val(str�걾��) = 0 And Len(str�걾��) > 0 Then
                                
                                strTmp�걾�� = str�걾��
                                str�걾�� = ""
                                For i = 1 To Len(strTmp�걾��)
                                    If IsNumeric(Mid(strTmp�걾��, i, 1)) Then
                                        str�걾�� = str�걾�� & Mid(strTmp�걾��, i, 1)
                                    End If
                                Next
                                str�걾�� = 1000 + Val(str�걾��)
                                str�ʿ� = "1"
                            End If
                            strSegment = ""
                        End If
                    Loop
                ElseIf strSegName = "PatTime" Then                         '����
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '�ı����ݣ�����ֱ����strsegment������
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '��ʽ����,������
                            strSegment = ""
                        End If

                        
                        If strFD_Name = "TestTime" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str���� = Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str���� = Mid_bin(strSegment, 1)
                            End If
                            If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd")
                            strSegment = ""
                        End If
                    Loop
                    
                ElseIf strSegName = "RBCHisto" Then                     'rbcֱ��ͼ
                    '����������
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- ���ݳ���
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- ȡͼ������
                        lngImgStart = Instr_Bin(strData, "RHistoData")
                        If lngImgStart > 0 Then
                            lngImgStart = lngImgStart + 11
                            strImgRBC = Graph_Text(Mid_bin(strData, lngImgStart, lngImgDataLen, False))
                            strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                            strSegment = ""
                        End If
                    Loop
                ElseIf strSegName = "PLTHisto" Then                     'pltֱ��ͼ
                    '����������
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- ���ݳ���
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- ȡͼ������
                        lngImgStart = Instr_Bin(strData, "PHistoData")
                        If lngImgStart > 0 Then
                            lngImgStart = lngImgStart + 11
                            strImgPLT = Graph_Text(Mid_bin(strData, lngImgStart, lngImgDataLen, False))
                            strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                            strSegment = ""
                        End If
                    Loop
'                ElseIf strSegName = "SepLine" Then                      'ֱ��ͼ������
                ElseIf strSegName = "BASOType" Or strSegName = "DIFFType" Then
                    '����������
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- ���ݳ���
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        
                        If strSegName = "BASOType" Then
                            lngImgStart = Instr_Bin(strData, "BASOTypeData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 13
                                Dim strBASOType As String
                                strBASOType = Mid_bin(strData, lngImgStart, lngImgDataLen, False)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFTypeData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 13
                                Dim strDiffType As String
                                strDiffType = Mid_bin(strData, lngImgStart, lngImgDataLen, False)
                                
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        
                        
                        End If
                    Loop
                ElseIf strSegName = "BASO" Or strSegName = "DIFF" Then  'ɢ��ͼ
                    '����������
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- ���ݳ���
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- ȡͼ������
                        If strSegName = "BASO" Then
                            lngImgStart = Instr_Bin(strData, "BASOData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 9
                                strImgBASO = DrawBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "BASO", str�걾��, strBASOType)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 9
                                strImgDiff = DrawBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "DIFF", str�걾��, strDiffType)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        End If
                    Loop
                ElseIf strSegName = "BASOBMP" Or strSegName = "DIFFBMP" Then
                    '����������
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- ���ݳ���
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- ȡͼ������
                        If strSegName = "BASOBMP" Then
                            lngImgStart = Instr_Bin(strData, "BASOBmpData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 12
                                strImgBASO = SaveBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "BASO", str�걾��)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFBmpData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 12
                                strImgBASO = SaveBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "DIFF", str�걾��)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        End If
                    Loop
                Else
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If
                End If
                
            Else
                If lngE > 0 Then
                    strData = Mid_bin(strData, lngE + 1, , False)
                Else
                    strData = ""
                End If
            End If
        Loop
        
        If Val(str�걾��) > 0 And IsDate(str����) And str��Ŀ <> "" Then
            strResult = str���� & "|" & Val(str�걾��) & "| |��ָѪ|" & str�ʿ� & str��Ŀ
            strImgData = ""
            If strImgDiff <> "" Then strImgData = strImgData & "^DIFF;107;1;" & strImgDiff
            If strImgBASO <> "" Then strImgData = strImgData & "^BASO;107;1;" & strImgBASO
            If strImgRBC <> "" Then strImgData = strImgData & "^RBC;0;260,300,0-0,50-50,100-100,150-150,200-200,250-250,300-fL" & strImgRBC
            If strImgPLT <> "" Then strImgData = strImgData & "^PLT;0;260,150,0-0,21.333-5,42.66-10,63.99-15,85.32-20,106.65-25,127.98-fL" & strImgPLT
            If strImgData <> "" Then strResult = strResult & vbCrLf & Mid(strImgData, 2)
        End If
        
    End If
    
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

Private Function DrawBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str�걾�� As String, ByVal strDataType As String) As String
    Dim strBinData As String
    Dim lngCount As Long
    Dim strfilename As String
    
    Dim frmPic As New frmGraph
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    If frmPic.Draw_Bc5500(strBinData, strfilename, strDataType) Then
        DrawBmpToFile = strfilename
    End If
End Function


Private Function SaveBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str�걾�� As String) As String
    '�������������ΪBMPͼ���ļ�
    Dim strfilename As String, lngFileNum As Long, lngloop As Long
    Dim strByte As String, ByteData() As Byte
    Dim strBinData As String
    Dim lngCount As Long
    
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    ReDim ByteData(lngCount)
    
    For lngloop = 0 To lngCount - 1
        strByte = Left(strBinData, 3)
        strBinData = Mid(strBinData, 4)
        ByteData(lngloop) = CByte("&H" & Mid(strByte, 2))
    Next
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strfilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBmpToFile = strfilename
End Function

Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Mid(Mid_bin(str_bin, lng_i, 1, False), 2)
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function