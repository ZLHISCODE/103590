Attribute VB_Name = "mdlArchive"
Option Explicit
'����
Public gblnCheck As Boolean '����CheckBox���ܹ�ѡ
Public gOldwinproc As Long 'ԭʼ��Ϣ���
Public Enum GeneralCtrlArchive
    '������Ϣ
    GCA_���ʽ = 0
    GCA_�������� = 1
    GCA_סԺ���� = 2
    GCA_������ = 3
    GCA_���� = 4
    GCA_�Ա� = 5
    GCA_�������� = 6
    GCA_���� = 7
    GCA_���� = 8
    GCA_���� = 9
    GCA_��� = 10
    GCA_������������ = 11
    GCA_���������� = 12
    GCA_��������Ժ���� = 13
    GCA_�����ص� = 14
    GCA_���� = 15
    GCA_���� = 16
    GCA_���֤�� = 17
    GCA_ְҵ = 18
    GCA_���� = 19
    GCA_��ͥ��ַ = 20
    GCA_��ͥ�绰 = 21
    GCA_��ͥ�ʱ� = 22
    GCA_���ڵ�ַ = 23
    GCA_�����ʱ� = 24
    GCA_��λ��ַ = 25
    GCA_��λ�绰 = 26
    GCA_��λ�ʱ� = 27
    GCA_��ϵ������ = 28
    GCA_��ϵ�˹�ϵ = 29
    GCA_��ϵ�˵绰 = 30
    GCA_��ϵ�˵�ַ = 31
    GCA_���� = 32
    GCA_��Ժ;�� = 33
    GCA_��Ժʱ�� = 34
    GCA_��Ժ���� = 35
    GCA_��Ժ���� = 36
    GCA_ת��1 = 38
    GCA_ת��2 = 39
    GCA_ת��3 = 40
    GCA_��Ժʱ�� = 41
    GCA_��Ժ���� = 42
    GCA_��Ժ���� = 43
    GCA_סԺ���� = 44
    '��ҽ���
    GCA_��Ժ��� = 45
    GCA_ȷ������ = 46
    GCA_����� = 47
    GCA_�ֻ��̶� = 48
    GCA_���������� = 49
    GCA_�����벡�� = 50
    GCA_�������ԺXY = 51
    GCA_��Ժ���ԺXY = 52
    GCA_��������Ժ = 53
    GCA_�ٴ��벡�� = 54
    GCA_�ٴ���ʬ�� = 55
    GCA_��ǰ������ = 56
    GCA_����ʱ�� = 57
    GCA_����ԭ�� = 58
    GCA_ҽԺ��Ⱦ��ԭѧ��� = 59
    GCA_���ȴ��� = 60
    GCA_�ɹ����� = 61
    GCA_����ԭ�� = 62
    '��ҽ���
    GCA_�������ԺZY = 63
    GCA_��Ժ���ԺZY = 64
    GCA_��֤ = 65
    GCA_�η� = 66
    GCA_��ҩ = 67
    GCA_������� = 68
    GCA_���ȷ��� = 69
    GCA_������ҩ = 70
    GCA_��ҽ�豸 = 71
    GCA_��ҽ���� = 72
    GCA_��֤ʩ�� = 73
    'סԺ���
    GCA_�������� = 74
    GCA_HBsAg = 75
    GCA_��Ѫǰ9���� = 76
    GCA_Ѫ�� = 77 '������Ϣ MZ
    GCA_HCVAb = 78
    GCA_����ʱ�� = 79 '������Ϣ MZ
    GCA_RH = 80 '������Ϣ MZ
    GCA_HIVAb = 81
    GCA_����״�� = 82 '������Ϣ MZ
    GCA_��Һ��Ӧ = 83
    GCA_��Ѫ��Ӧ = 84
    GCA_���ϸ�� = 85
    GCA_��ѪС�� = 86
    GCA_��Ѫ�� = 87
    GCA_��ȫѪ = 88
    GCA_������ = 89
    GCA_������� = 90
    GCA_ҽѧ��ʾ = 91 '������Ϣ MZ
    GCA_����ҽѧ��ʾ = 92 '������Ϣ MZ
    GCA_��Ժ��ʽ = 93
    GCA_ת����� = 94
    GCA_��Ժǰ�� = 95
    GCA_��ԺǰСʱ = 96
    GCA_��Ժǰ���� = 97
    GCA_��Ժ���� = 98
    GCA_��Ժ��Сʱ = 99
    GCA_��Ժ����� = 100
    GCA_����Ժ���� = 101
    GCA_31��Ŀ�� = 102
    GCA_������Сʱ = 103
    GCA_�������� = 104
    GCA_����ҽʦ = 105
    GCA_������ = 106
    GCA_����ҽʦ = 107
    GCA_����ҽʦ = 108
    GCA_סԺҽʦ = 109
    GCA_����ҽʦ = 110
    GCA_�о���ҽʦ = 111
    GCA_����ҽʦ = 111
    GCA_ʵϰҽʦ = 112
    GCA_�ʿ�ҽʦ = 113
    GCA_���λ�ʿ = 114
    GCA_�ʿػ�ʿ = 115
    GCA_�ʿ����� = 116
    GCA_�������� = 117
    '����
    GCA_ѹ�������ڼ� = 118 '��ҳ1 YN
    GCA_ѹ������ = 119 '��ҳ1 YN
    GCA_������׹���˺� = 120 '��ҳ1 YN
    GCA_������׹��ԭ�� = 121 '��ҳ1 YN
    GCA_��֢�໤ = 123 'HN
    GCA_��֢�໤���� = 124 'HN
    GCA_��֢�໤Сʱ = 125 'HN
    GCA_�������� = 126 'HN
    GCA_��������T = 127 'HN
    GCA_��������M = 128 'HN
    GCA_��������N = 129 'HN
    GCA_Apgar = 130 'HN
    GCA_�ٴ�·������ = 131 'HN
    GCA_��Ⱦ�� = 132 'HN
    GCA_DrGs���� = 133 'HN
    GCA_����ҩ�� = 134 'SC
    GCA_�ٴ����� = 135 'SC
    GCA_��Ժ͸�����ص�ֵ = 136 'SC
    '������Ϣ ST HN
    GCA_����֤�� = 122
    GCA_Email = 137
    GCA_QQ = 138
    '��ҽ���  ST
    GCA_��Ⱦ��λ = 145
    GCA_��Ⱦ������ = 146
    'סԺ��� SC
    GCA_��׵��� = 139
    GCA_�˳�ԭ�� = 140 '��ҳ1 YN
    GCA_����ԭ�� = 141 '��ҳ1 YN
    GCA_Ժ�ڻ������ = 142
    GCA_��Ժ������� = 143
    GCA_����������� = 144
    '��ҳ1
    GCA_��֢�໤�� = 147
    GCA_�ط����ʱ�� = 148
    GCA_Լ����ʱ�� = 149
    GCA_Լ����ʽ = 150
    GCA_Լ������ = 151
    GCA_Լ��ԭ�� = 152
    GCA_��������Ժ��ʽ = 153
    '������Ϣ MZ
    GCA_����� = 154
    GCA_�໤�� = 155
    GCA_�Ļ��̶� = 156
    GCA_����ժҪ = 157
    GCA_ȥ�� = 158
    GCA_������ַ = 159
    GCA_���� = 160
    GCA_���� = 161
    GCA_���� = 162
    GCA_Ѫѹ = 163
    '������Ϣ ZY
    GCA_סԺ�� = 164
    GCA_��Ժת�� = 165
    GCA_������ϵ = 166
    'SC:��ҳ2
    GCA_����һ��ס��Ժʱ�� = 167
    'סԺ���
    GCA_��������ʬ�� = 168
    GCA_�໤�����֤�� = 172
End Enum

Public Enum CheckCtrlArchive
    '������Ϣ
    CHKA_����Ժ = 0
    CHKA_��Ժǰ����Ժ���� = 1
    '��ҽ���
    CHKA_�Ƿ�ȷ�� = 2
    CHKA_ҽԺ��Ⱦ����ԭѧ��� = 3
'    CHKA_��������ʬ�� = 4
    CHKA_�·����� = 5
    '��ҽ���
    CHKA_Σ�� = 6
    CHKA_��֢ = 7
    CHKA_���� = 8
    'סԺ���
    CHKA_���Ѳ��� = 9 '��ҽ��� SC
    CHKA_ʾ�̲��� = 10
    CHKA_���в��� = 11
    CHKA_���� = 12
    '����
    CHKA_CT = 13
    CHKA_MRI = 14
    CHKA_��ɫ������ = 15
    CHKA_ϸ�������걾�ͼ� = 16 'HN
    CHKA_������ = 17 'HN
    '��ҽ��� SC
    CHKA_סԺ�ڼ�没�ػ�Σ = 18 '��ҳ1 YN
    'סԺ��� SC
    CHKA_����·�� = 19 '��ҳ1 YN
    CHKA_���·�� = 20 '��ҳ1 YN
    CHKA_���� = 21 '��ҳ1 YN
    CHKA_������� = 22
    '���� SC
    CHKA_סԺ�ڼ�����Լ�� = 23 'YN סԺ�ڼ�ʹ������Լ��
    '���������� YN
    CHKA_Χ�������� = 24
    CHKA_������� = 25
    '��ҳ1 YN
    CHKA_�˹������ѳ� = 26
    CHKA_�ط���֢ҽѧ�� = 27
    'MZ ������Ϣ
    CHKA_���� = 28
    CHKA_��Ⱦ���ϴ� = 29
    'SC:��ҳ2
    CHKA_�Ƿ���ͬһ���� = 30
    '������Ϣ��OM)
    CHKA_�޹�����¼ = 31
End Enum

Public Function ArchivezlRefresh() As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    On Error GoTo errH
    Call ClearPageContent
    If gclsPros.����ID <> 0 Then
        Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.MedPageSandard = ST_������ҳ, "NULL", "")) '������ҳ�Լ�������Ϣ
        If gclsPros.PatiInfo.EOF Then Exit Function
        If gclsPros.MedPageSandard = ST_������ҳ Then
            gclsPros.��Ժ����ID = Val(gclsPros.PatiInfo!����id & "")
        Else
            gclsPros.��Ժ����ID = Val(gclsPros.PatiInfo!��Ժ����ID & "")
        End If
        If Not ArchiveInitEnv Then Exit Function
        Call ArchiveLoadPageData(gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.MedPageSandard = ST_������ҳ, "NULL", ""))
    End If
    Call ArchiveSetPageHeight
    Call ArchiveFormResize
    ArchivezlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ArchiveInitEnv() As Boolean
'���ܣ����ý��棬����ʼ����
'         blnFormLoad�Ƿ���FormLoad����
    '�������
    If Not InitTableAller Then Exit Function
    If Not InitTableDiag Then Exit Function
    If gclsPros.MedPageSandard <> ST_������ҳ Then
        If Not InitTableOPS Then Exit Function
        If Not InitTableKSS Then Exit Function
        If Not InitTableFlxAddICU Then Exit Function
        If Not InitTablefMain Then Exit Function
        If gclsPros.ReadPages Then
            If Not InitTableSpirit Then Exit Function
            If Not InitTableChemoth Then Exit Function
            If Not InitTableRadioth Then Exit Function
        End If
        If Not InitTableICUInstruments Then Exit Function
        If Not InitTableInfect Then Exit Function
        If Not InitTableSample Then Exit Function
        If Not InitTableTSJC Then Exit Function
    End If
    ArchiveInitEnv = True
End Function

Private Function ArchiveLoadPageData(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long = 1, Optional ByVal str�Һŵ� As String) As Boolean
'���ܣ����Ӳ����������ݼ���
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrTmp As Variant
    On Error GoTo errH
    gblnCheck = True
    With gclsPros.CurrentForm
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.����ID, gclsPros.��ҳID, str�Һŵ�) '�ӱ���Ϣ
        '���ز�����Ϣ
        If Not gclsPros.PatiInfo.EOF Then
            For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
                 Call ArchiveSetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "")
            Next
        End If
        '������ҳ�ӱ�������Ϣ�ӱ����
        If Not gclsPros.AuxiInfo.EOF Then
            gclsPros.AuxiInfo.MoveFirst
            For i = 1 To gclsPros.AuxiInfo.RecordCount
                 Call ArchiveSetCtrlValues(gclsPros.AuxiInfo!��Ϣ�� & "", gclsPros.AuxiInfo!��Ϣֵ & "", gclsPros.AuxiInfo!���� & "")
                 gclsPros.AuxiInfo.MoveNext
            Next
        End If
        '����������Ϣ����
        If gclsPros.MedPageSandard = ST_������ҳ Then
            'Ѫѹ��Ϣ��Ҫ������װ
            If .txtInfo(GCA_Ѫѹ).Tag Like "*|*" Then
                arrTmp = Split(.txtInfo(GCA_Ѫѹ).Tag, "|")
                .txtInfo(GCA_Ѫѹ).Text = arrTmp(0) & " / " & arrTmp(1) & " " & IIf(.lblInfo(GCA_Ѫѹ).Tag = "", "mmHg", .lblInfo(GCA_Ѫѹ).Tag)
            ElseIf .txtInfo(GCA_Ѫѹ).Tag <> "" Then
                .txtInfo(GCA_Ѫѹ).Text = " / " & .txtInfo(GCA_Ѫѹ).Tag & " " & IIf(.lblInfo(GCA_Ѫѹ).Tag = "", "mmHg", .lblInfo(GCA_Ѫѹ).Tag)
            End If
            .txtInfo(GCA_Ѫѹ).Tag = "": .lblInfo(GCA_Ѫѹ).Tag = ""
             '��Ϊ66029���ⲿ�����ݴ洢λ�÷����仯�����¶�ȡ��������
            Set rsTmp = GetCareData(gclsPros.����ID, gclsPros.��ҳID)
            rsTmp.Filter = "��Ϣ��='���'"
            If Not rsTmp.EOF Then .txtInfo(GCA_���).Text = rsTmp!��Ϣֵ & " " & rsTmp!��λ
            rsTmp.Filter = "��Ϣ��='����'"
            If Not rsTmp.EOF Then .txtInfo(GCA_����).Text = rsTmp!��Ϣֵ & " " & rsTmp!��λ
            rsTmp.Filter = "��Ϣ��='����'"
            If Not rsTmp.EOF Then .txtInfo(GCA_����).Text = rsTmp!��Ϣֵ & " " & rsTmp!��λ
            rsTmp.Filter = "��Ϣ��='����ѹ'"
            If Not rsTmp.EOF Then .txtInfo(GCA_Ѫѹ).Text = IIf(NVL(rsTmp!��Ϣֵ) = "", "   ", NVL(rsTmp!��Ϣֵ))
            rsTmp.Filter = "��Ϣ��='����ѹ'"
            If Not rsTmp.EOF Then .txtInfo(GCA_Ѫѹ).Text = .txtInfo(GCA_Ѫѹ).Text & " / " & IIf(NVL(rsTmp!��Ϣֵ) = "", "   ", NVL(rsTmp!��Ϣֵ)) & " " & rsTmp!��λ
            rsTmp.Filter = "��Ϣ��='����'"
            If Not rsTmp.EOF Then .txtInfo(GCA_����).Text = rsTmp!��Ϣֵ & IIf(rsTmp!��Ϣֵ & "" = "", "", " ��/��")
            rsTmp.Filter = "��Ϣ��='����'"
            If Not rsTmp.EOF Then .txtInfo(GCA_����).Text = rsTmp!��Ϣֵ & IIf(rsTmp!��Ϣֵ & "" = "", "", " ��/��")
        Else
            '���۲�����סԺ��
            If Val(gclsPros.PatiInfo!�������� & "") <> 0 Then
                .lblInfo(GCA_��������).Visible = False
                .txtInfo(GCA_��������).Visible = False
                .lblInfo(GCA_סԺ��).Visible = False
                .txtInfo(GCA_סԺ��).Visible = False
            End If
            'סԺ��������
            If Not IsNull(gclsPros.PatiInfo!��Ժ����) Then
                .txtInfo(GCA_סԺ����).Text = DateDiff("d", gclsPros.PatiInfo!��Ժ����, gclsPros.PatiInfo!��Ժ����)
            Else
                .txtInfo(GCA_סԺ����).Text = DateDiff("d", gclsPros.PatiInfo!��Ժ����, zlDatabase.Currentdate)
            End If
            If Val(.txtInfo(GCA_סԺ����).Text) = 0 Then .txtInfo(GCA_סԺ����).Text = "1"

            '�Զ���ȡת�ƿ��Ҽ��������(�����)
            '---------------------------------------------------------------
            If .txtInfo(GCA_ת��1).Text = "" And .txtInfo(GCA_ת��2).Text = "" And .txtInfo(GCA_ת��3).Text = "" Then
                Set rsTmp = GetPatiTransfer(gclsPros.����ID, gclsPros.��ҳID)
                For i = 1 To rsTmp.RecordCount
                    If i = 1 Then
                        .txtInfo(GCA_ת��1).Text = rsTmp!��������
                    ElseIf i = 2 Then
                        .txtInfo(GCA_ת��2).Text = rsTmp!��������
                    ElseIf i = 3 Then
                        .txtInfo(GCA_ת��3).Text = rsTmp!��������
                        Exit For
                    End If
                    rsTmp.MoveNext
                Next
            End If
            If .txtInfo(GCA_��Ժ����).Text = "" Or .txtInfo(GCA_��Ժ����).Text = "" Then
                Set rsTmp = GetPatiRoom(gclsPros.����ID, gclsPros.��ҳID)
                If .txtInfo(GCA_��Ժ����).Text = "" Then .txtInfo(GCA_��Ժ����).Text = rsTmp!��Ժ���� & ""
                If .txtInfo(GCA_��Ժ����).Text = "" Then .txtInfo(GCA_��Ժ����).Text = rsTmp!��Ժ���� & ""
            End If
        End If
        '��ȡ���
        Set rsTmp = GetPatiDiagData(gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.MedPageSandard = ST_������ҳ, 0, 1), , , gclsPros.Moved)
        If gclsPros.MedPageSandard <> ST_������ҳ Then
            '���ز�ԭѧ���
            Call FilterDiagByType(rsTmp, DT_��ԭѧ���)
            If Not rsTmp.EOF Then
                .txtInfo(GCA_ҽԺ��Ⱦ��ԭѧ���).Text = rsTmp!������� & ""
            End If
        End If
        '�������
        Call ArchiveLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.MedPageSandard <> ST_������ҳ, "1,2,3,5,6,7,10", "1"))
        If gclsPros.Have��ҽ Then
            Call ArchiveLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.MedPageSandard <> ST_������ҳ, "11,12,13", "11"))
        End If
        '������Ϣ����
        If gclsPros.MedPageSandard <> ST_������ҳ Then
            Call ArchiveLoadAller(.vsAller, GetAllerData(gclsPros.����ID, gclsPros.��ҳID))
        ElseIf .chkInfo(CHKA_�޹�����¼).Value = 0 Then '��ѡ�޹�����¼���򲻼��ع�����¼
            Call ArchiveLoadAller(.vsAller, GetAllerData(gclsPros.����ID, gclsPros.��ҳID))
        End If
        If gclsPros.MedPageSandard <> ST_������ҳ Then
            '����������Ϣ
            Call ArchiveLoadOPS(.vsOPS, GetOPSData(gclsPros.����ID, gclsPros.��ҳID, , gclsPros.Moved))
            '��Ϸ����������
            Set rsTmp = GetDiagMatchData(gclsPros.����ID, gclsPros.��ҳID)
            Do While Not rsTmp.EOF
                Select Case rsTmp!��������
                    Case 1 '�������Ժ
                        .txtInfo(GCA_�������ԺXY).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 2 '��Ժ���Ժ
                        .txtInfo(GCA_��Ժ���ԺXY).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 3 '�����벡��
                        .txtInfo(GCA_�����벡��).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 4 '�ٴ��벡��
                        .txtInfo(GCA_�ٴ��벡��).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 5 '�ٴ���ʬ��
                        .txtInfo(GCA_�ٴ���ʬ��).Text = decode(NVL(rsTmp!�������, 0), 0, "δ��", 1, "����", 2, "������", 3, "���϶�", "-")
                    Case 6 '��ǰ������
                        .txtInfo(GCA_��ǰ������).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 7 '��������Ժ
                        .txtInfo(GCA_��������Ժ).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 11 '��ҽ�������Ժ
                        .txtInfo(GCA_�������ԺZY).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 12 '��ҽ��Ժ���Ժ
                        .txtInfo(GCA_��Ժ���ԺZY).Text = decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
                    Case 13 '��ҽ��֤
                        .txtInfo(GCA_��֤).Text = decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
                    Case 14 '��ҽ�η�
                        .txtInfo(GCA_�η�).Text = decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
                    Case 15 '��ҽ��ҩ
                        .txtInfo(GCA_��ҩ).Text = decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
                End Select
                rsTmp.MoveNext
            Loop
            '����ҩ��
            Call ArchiveLoadKSS(.vsKSS, GetKSSData(gclsPros.����ID, gclsPros.��ҳID))
            '������Ŀ
            If gclsPros.ReadPages Then
                Call ArchiveLoadPageMedRec(gclsPros.����ID, gclsPros.��ҳID)
            End If
            Call ArchiveLoadOtherInfo(gclsPros.����ID, gclsPros.��ҳID)
        End If
    End With
    gblnCheck = False
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ArchiveSetPageHeight()
'���ܣ�����ҳ��������չ��״̬���ý���ߴ�
'˵����Tag=1��ʾ����
    Dim i As Long, intCurIdx As Integer
    With gclsPros.CurrentForm
        For i = 0 To .fraMain.UBound
            If Val(.picSize(i).Tag) = 0 Then
                .fraMain(i).Height = Val(.fraMain(i).Tag)
                Set .picSize(i).Picture = .imgSize.ListImages("-").Picture
            Else
                .fraMain(i).Height = 225
                Set .picSize(i).Picture = .imgSize.ListImages("+").Picture
            End If
        Next
        
        intCurIdx = 0
        For i = 1 To .fraMain.UBound
            If .fraMain(i).Enabled Then
                .fraMain(i).Top = .fraMain(intCurIdx).Top + .fraMain(intCurIdx).Height + 100
                intCurIdx = i
            End If
        Next
       .fraBack.Height = .fraMain(intCurIdx).Top + .fraMain(intCurIdx).Height + .fraMain(0).Top
        Call ArchiveSetScrollbar
    End With
End Sub

Public Sub ArchiveSetScrollbar()
'���ܣ����ݵ�ǰ����ߴ����ù������ɼ��Լ��������
    With gclsPros.CurrentForm
        If .fraBack.Width + IIf(.vsc.Visible, .vsc.Width, 0) <= .picBack.ScaleWidth Then
            .hsc.Visible = False
        Else
            .hsc.Min = 0
            .hsc.SmallChange = 5
            .hsc.LargeChange = 50
            If Not .hsc.Visible Then .hsc.Value = 0
            .hsc.Visible = True
        End If
        
        If .fraBack.Height + IIf(.hsc.Visible, .hsc.Height, 0) <= .picBack.ScaleHeight Then
            .vsc.Visible = False
        Else
            .vsc.Min = 0
            .vsc.SmallChange = 5
            .vsc.LargeChange = 50
            If Not .vsc.Visible Then .vsc.Value = 0
            .vsc.Visible = True
        End If
        .hsc.Max = (.picBack.ScaleWidth - .fraBack.Width - IIf(.vsc.Visible, .vsc.Width, 0)) / Screen.TwipsPerPixelX
        .vsc.Max = (.picBack.ScaleHeight - .fraBack.Height - IIf(.hsc.Visible, .hsc.Height, 0)) / Screen.TwipsPerPixelY
        .fraVH.Visible = .vsc.Visible And .hsc.Visible
    End With
End Sub

Public Sub ArchiveFormResize()
    On Error Resume Next
    With gclsPros.CurrentForm
        .picBack.Left = 0
        .picBack.Top = 0
        .picBack.Width = .ScaleWidth
        .picBack.Height = .ScaleHeight
        .hsc.Left = 0
        .hsc.Top = .picBack.ScaleHeight - .hsc.Height
        .hsc.Width = .picBack.ScaleWidth - IIf(.vsc.Visible, .vsc.Width, 0)
        .vsc.Top = 0
        .vsc.Left = .picBack.ScaleWidth - .vsc.Width
        .vsc.Height = .picBack.ScaleHeight - IIf(.hsc.Visible, .hsc.Height, 0)
        If .fraVH.Visible Then
            .fraVH.Left = .vsc.Left
            .fraVH.Top = .hsc.Top
            .fraVH.Refresh
        End If
        Call ArchiveSetScrollbar
    End With
End Sub

Public Sub ArchiveFormKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = gclsPros.CurrentForm.vsc.Value
    lngMin = gclsPros.CurrentForm.vsc.Min
    lngMax = gclsPros.CurrentForm.vsc.Max
    If lngMax <= lngMin Then '��ֱ������δ����
        If intKeyCode = vbKeyPageDown Then '��
            If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
                gclsPros.CurrentForm.vsc.Value = lngCur + (lngMax - lngMin) / 10
            Else
                gclsPros.CurrentForm.vsc.Value = lngMax
            End If
        Else '��
            If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
                gclsPros.CurrentForm.vsc.Value = lngCur - (lngMax - lngMin) / 10
            Else
                gclsPros.CurrentForm.vsc.Value = lngMin
            End If
        End If
    End If
End Sub

Public Function ArchiveFormLoad() As Boolean
    With gclsPros.CurrentForm
        '�������ߴ�
        .vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
        .hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
        .fraVH.Width = .vsc.Width: .fraVH.Height = .hsc.Height
        .fraBack.Left = 0: .fraBack.Top = 0
        .picBack.BackColor = .fraBack.BackColor
    End With
End Function

Public Sub ArchivepicSizeClick(ByRef intIndex As Integer)
    With gclsPros.CurrentForm
        .picSize(intIndex).Tag = IIf(Val(.picSize(intIndex).Tag) = 0, 1, 0)
        Call ArchiveSetPageHeight
        Call ArchiveFormResize
        If Not .vsc.Visible Then .fraBack.Top = 0
        If Not .hsc.Visible Then .fraBack.Left = 0
    End With
End Sub

Public Sub ArchivechkInfoClick(ByRef intIndex As Integer)
    If Not gblnCheck Then
        gblnCheck = True
        gclsPros.CurrentForm.chkInfo(intIndex).Value = IIf(gclsPros.CurrentForm.chkInfo(intIndex).Value = 1, 0, 1)
        gblnCheck = False
    End If
End Sub

Public Function ArchiveSetCtrlValues(ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal str���ӱ��� As String) As Boolean
'���ܣ����ÿؼ�ֵ
'����  strInfoName=��Ϣ��
'      strInfoValue=��Ϣֵ
'      str���ӱ���=����������Ŀ�����ж�
    Dim str�ؼ��� As String
    Dim lngCount As Long, i As Long, j As Long, LngRow As Long
    Dim arrTmp As Variant, strTmp As String
    Dim vsTmp As VSFlexGrid, lstTmp As ListBox
    Dim intIndex As Integer, intIndexTmp As Integer
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Select Case strInfoName
        Case "������4", "������5", "������6"
            gclsPros.MainInfoRec.Filter = "��Ϣ��='������'"
        Case "CT", "PETCT", "˫ԴCT", "XƬ", "B��", "�����Ķ�ͼ", "MRI", "ͬλ�ؼ��"
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                gclsPros.MainInfoRec.Filter = "��Ϣ��='������'"
            Else
                gclsPros.MainInfoRec.Filter = "��Ϣ��='" & strInfoName & "'"
            End If
        Case "����ѹ", "����ѹ", "Ѫѹ��λ"
            gclsPros.MainInfoRec.Filter = "��Ϣ��='Ѫѹ'"
        Case Else
            gclsPros.MainInfoRec.Filter = "��Ϣ��='" & strInfoName & "'"
    End Select
    '��Ϣδ�ڼ�¼����ע�ᣬ���������ݼ�����չ���ͣ��磺����������Ŀ���Ͽ����ؼ�¼
    If gclsPros.MainInfoRec.EOF Then
        If strInfoValue = "" Then Exit Function
        '�������������,�ϰ濹������Ϣ
        If strInfoName Like "������*" And IsNumeric(Mid(strInfoName, 4)) Then
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                LngRow = -1
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_����ҩ����) = "" Then LngRow = i: Exit For
                Next
                If i > .Rows - 1 Then .AddItem "": LngRow = i
                '���������ݣ�����ҳ�ӱ����ȶ�����
                .RowData(LngRow) = GetKSSID(strInfoValue)
                If Val(.RowData(LngRow) & "") <> 0 Then
                    .TextMatrix(LngRow, KI_����ҩ����) = strInfoValue
                    .Cell(flexcpData, LngRow, KI_����ҩ����) = .TextMatrix(LngRow, KI_����ҩ����)
                End If
                Call SetKSSSerial
            End With
        '������Ŀ�������������ԭ�����ڵĲ�����ҳ�ӱ���Ϣ�������Ƴ�ͻ��
        '������Ҳ��������ϵĴӱ���Ϣʱ�ż��ز���������Ŀ
        ElseIf str���ӱ��� <> "" Then
            Set vsTmp = gclsPros.CurrentForm.vsfMain
            With vsTmp
                For i = 0 To 3 Step 3
                    LngRow = -1: LngRow = .FindRow(strInfoName, , i)
                    If LngRow >= 0 Then
                        If .TextMatrix(LngRow, i + 2) = "�Ƿ�" Then
                            .Cell(flexcpChecked, LngRow, i + 1) = IIf(Val(strInfoValue) = 0, 2, 1)
                            Exit For
                        Else
                            .TextMatrix(LngRow, i + 1) = strInfoValue
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
    Else
        str�ؼ��� = gclsPros.MainInfoRec!�ؼ��� & ""
        With gclsPros.CurrentForm
            '������Ϣ��չ״̬
            If gclsPros.MainInfoRec!ExpState = 0 Then
                intIndex = Val(gclsPros.MainInfoRec!Index & "")
                Select Case str�ؼ���
                    Case "txtInfo"
                        .txtInfo(intIndex).Text = strInfoValue
                        Select Case intIndex
                            Case GCA_����ʱ��, GCA_��������
                                strInfoValue = Format(strInfoValue, IIf(Format(strInfoValue, "HH:mm") <> "00:00", "yyyy-MM-dd HH:mm", "yyyy-MM-dd"))
                            Case GCA_��Ժʱ��
                                strInfoValue = Format(strInfoValue, "yyyy-MM-dd HH:mm")
                            Case GCA_���, GCA_����
                                If gclsPros.MedPageSandard = ST_������ҳ And strInfoValue <> "" Then strInfoValue = strInfoValue & " " & decode(intIndex, GCA_���, "cm", GCA_����, "Kg")
                            Case GCA_����, GCA_����, GCA_����, GCA_����������, GCA_��������Ժ����
                                 If strInfoValue <> "" Then strInfoValue = strInfoValue & " " & decode(intIndex, GCA_����, "��", GCA_����, "��/��", GCA_����, "��/��", GCA_����������, "��", GCA_��������Ժ����, "��")
                            Case GCA_����Ժ����
                                .lblInfo(intIndex).Caption = "��Ժ" & IIf(Val(strInfoValue & "") = 0, 31, 7) & "��������Ժ�ƻ�"
                            Case GCA_31��Ŀ��
                                .txtInfo(GCA_����Ժ����).Text = IIf(strInfoValue <> "", "��", "��")
                            Case GCA_��֢�໤����, GCA_��֢�໤Сʱ
                                If strInfoValue <> "" Then .txtInfo(GCA_��֢�໤) = "��"
                            Case GCA_��������
                                If strInfoValue <> "" Then strInfoValue = GetNameByCode("��������", strInfoValue)
                            Case GCA_ȷ������ '��Ϊȷ���־��ȷ������ǰ���أ���˿���������
                                If .chkInfo(CHKA_�Ƿ�ȷ��).Value = 0 Then
                                    strInfoValue = ""
                                Else
                                    strInfoValue = Format(strInfoValue, "yyyy-MM-dd HH:mm")
                                End If
                            Case GCA_�ɹ�����
                                If Val(.txtInfo(GCA_���ȴ���).Text) = 0 Then strInfoValue = ""
                            Case GCA_��������
                                If .chkInfo(CHKA_����).Value = 0 Then
                                    strInfoValue = ""
                                Else
                                    strInfoValue = IIf(Val(gclsPros.PatiInfo!�����־ & "") = 9, "", Val(strInfoValue & "")) & _
                                                decode(Val(gclsPros.PatiInfo!�����־ & ""), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����")
                                End If
                            Case GCA_��Ⱦ��λ
                                If strInfoValue <> "" Then
                                    Set rsTmp = GetBaseCode(strInfoName)
                                    strTmp = ""
                                    If InStr(strInfoValue, "|") > 0 Then
                                        strInfoValue = Replace(strInfoValue, "|", ",") '����|���ָ����ת��Ϊ����
                                    End If
                                    Set rsTmp = GetBaseCode(strInfoName)
                                    For i = 1 To rsTmp.RecordCount
                                        If InStr("," & strInfoValue & ",", "," & rsTmp!���� & ",") > 0 Then
                                            strTmp = strTmp & "," & NVL(rsTmp!����)
                                        End If
                                        rsTmp.MoveNext
                                    Next
                                    If strTmp <> "" Then
                                        strInfoValue = Mid(strInfoValue, 2)
                                    Else
                                        strInfoValue = ""
                                    End If
                                End If
                            Case GCA_�˳�ԭ��, GCA_����ԭ��
                                If intIndex = GCA_�˳�ԭ�� Then
                                    .chkInfo(CHKA_���·��).Value = IIf(strInfoValue = "1", 1, 0)
                                    If strInfoValue = "1" Then strInfoValue = ""
                                Else
                                    .chkInfo(CHKA_����).Value = IIf(strInfoValue <> "", 1, 0)
                                    If strInfoValue = "1" Then strInfoValue = ""
                                End If
                            Case GCA_�����������
                                .chkInfo(CHKA_�������).Value = IIf(strInfoValue <> "0", 1, 0)
                                If strInfoValue = "0" Then strInfoValue = ""
                            Case GCA_Ժ�ڻ������, GCA_��Ժ�������
                                .chkInfo(CHKA_�������).Value = 1
                            Case GCA_Ѫѹ
                                Select Case strInfoName
                                    Case "����ѹ"
                                        .txtInfo(intIndex).Tag = strInfoValue & "|" & .txtInfo(intIndex).Tag
                                    Case "����ѹ"
                                        .txtInfo(intIndex).Tag = .txtInfo(intIndex).Tag & strInfoValue
                                    Case "Ѫѹ��λ"
                                        .lblInfo(intIndex).Tag = strInfoValue
                                End Select
                            Case GCA_����״��
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 0, "δ����", 1, "����1̥", 2, "����2̥������", 4, "����")
                            Case GCA_��Ѫ��Ӧ
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 0, "��", 1, "��", 2, "δ��")
                            Case GCA_�ٴ�·������
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "δ����", 2, "�����˳�", 3, "���")
                            Case GCA_DrGs����
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "��", 2, "������", 3, "������", 4, "���߶���")
                            Case GCA_��Ⱦ��
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "����", 2, "����", 3, "����")
                            Case GCA_��������
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "0��", 2, "I��", 3, "����", 4, "����", 5, "����", 6, "����")
                            Case GCA_��������ʬ��
                                strInfoValue = IIf(strInfoValue = "1", "��", " ")
                                '����Ժ��ʽΪ����ʱ������ʬ����ֵ��չʾΪ��
                                If .txtInfo(intIndex).Tag = "1" And strInfoValue = "" Then
                                    strInfoValue = "��"
                                End If
                                If strInfoValue = "" Then .txtInfo(intIndex).Tag = "0"
                            Case GCA_��Ժ��ʽ
                                If strInfoValue = "����" And .txtInfo(GCA_��������ʬ��).Tag = "0" Then
                                    .txtInfo(GCA_��������ʬ��) = "��"
                                Else
                                    .txtInfo(GCA_��������ʬ��).Tag = "1"
                                End If
                            Case GCA_���֤��
                                If zlStr.ActualLen(strInfoValue) > 12 And gclsPros.IsMaskID Then   '�������֤������
                                    strInfoValue = Mid(strInfoValue, 1, 12) & String(Len(Mid(strInfoValue, 13, 2)), "*") & Mid(strInfoValue, 15)
                                End If
                        End Select
                        .txtInfo(intIndex).Text = strInfoValue
                    Case "chkInfo"
                        .chkInfo(intIndex).Value = IIf(Val(strInfoValue) = 0, 0, 1)
                    Case "lstInfection", "lstAdvEvent"
                        If strInfoName = "��Ⱦ����" Then
                            Set lstTmp = .lstInfection
                        ElseIf strInfoName = "�����¼�" Then
                            Set lstTmp = .lstAdvEvent
                        End If
                        If InStr(strInfoValue, "|") > 0 Then
                            strInfoValue = Replace(strInfoValue, "|", ",") '����|���ָ����ת��Ϊ����
                        End If
                        Set rsTmp = GetBaseCode(strInfoName)
                        For i = 1 To rsTmp.RecordCount
                            If InStr("," & strInfoValue & ",", "," & rsTmp!���� & ",") > 0 Then
                                lstTmp.AddItem NVL(rsTmp!����)
                            End If
                            rsTmp.MoveNext
                        Next
                End Select
            ElseIf gclsPros.MainInfoRec!ExpState = 1 Then
                If str�ؼ��� <> "vsTSJC" Then
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                End If
                Select Case strInfoName
                    Case "����ʱ��", "ת�Ƽ�¼"
                        '�����ʽ:��Ժǰ(�죬Сʱ,����)|��Ժ��(�죬Сʱ,����)
                        If strInfoName = "����ʱ��" Then
                            strTmp = Replace(strInfoValue, "|", ",")
                            strTmp = strTmp & ",,,,,"
                        Else
                            strTmp = strInfoValue & ",,,"
                        End If
                        arrTmp = Split(strTmp, ",")
                        For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                            .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                            gclsPros.SecdInfoRec.MoveNext
                        Next
                    Case Else
                        If str�ؼ��� = "vsTSJC" Then
                            If strInfoName Like "������*" And gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                                intIndex = Val(Mid(strInfoName, 5, 1)) - 4
                            Else
                                intIndex = decode(strInfoName, "CT", TR_CT, "PETCT", TR_PETCT, "˫ԴCT", TR_˫ԴCT, _
                                            "XƬ", TR_XƬ, "B��", TR_B��, "�����Ķ�ͼ", TR_�����Ķ�ͼ, "MRI", TR_MRI, "ͬλ�ؼ��", TR_ͬλ�ؼ��, -1)
                                strInfoValue = decode(Val(strInfoValue), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
                            End If
                            If intIndex <> -1 Then
                                .vsTSJC.TextMatrix(intIndex, 1) = strInfoValue
                                .vsTSJC.Cell(flexcpData, intIndex, 1) = strInfoValue
                            End If
                        End If
                End Select
            ElseIf gclsPros.MainInfoRec!ExpState = 2 Then
            '����ʱ����
            End If
        End With
    End If
    ArchiveSetCtrlValues = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ArchiveLoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String)
'���ܣ�����ϼ��ص����
'������vsDiagInput=��Ҫ������ϵı��
'      rsInput=��ȡ����ϼ�¼��
'      strDiagType=��������ַ������������Զ��ŷָ�
'˵����ArchiveLoadMedPageData���Ӻ���

    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long
    Dim bln�ֻ��̶� As Boolean
    Dim bln��ҽ As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String

    On Error GoTo errH
    With vsDiagInput
        bln��ҽ = vsDiagInput.Name = "vsDiagXY"
        '�������
        arrTmp = Split(strDiagType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            Call FilterDiagByType(rsInput, Val(arrTmp(i))) '�������
            Do While Not rsInput.EOF
                If rsInput!������� = 1 Then
                    'ȷ����ǰ��ʾ��
                    LngRow = .FindRow(arrTmp(i), , DI_��Ϸ���, , True)
                    For j = LngRow To .Rows - 1
                        If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(arrTmp(i)) Then
                            LngRow = j
                            If .TextMatrix(j, DI_�������) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    '������
                    If .TextMatrix(LngRow, DI_�������) <> "" Then
                        LngRow = LngRow + 1: .AddItem "", LngRow
                        .TextMatrix(LngRow, DI_��Ϸ���) = arrTmp(i)
                        If gclsPros.MedPageSandard = ST_������ҳ Then .TextMatrix(LngRow, DI_�������) = IIf(Val(arrTmp(i)) = DT_�������XY, "��ҽ", "��ҽ")
                    End If
                    
                    strTmp = rsInput!������� & ""
                    '��ȡ��ϱ��룬�������Ϊ(����)��������(����)����(֤��) ���͵Ŀ��Ի�ȡ�������
                    If strTmp Like "(?*)?*" Then
                        lngPos = InStr(1, strTmp, ")")
                        .TextMatrix(LngRow, DI_��ϱ���) = Mid(strTmp, 2, lngPos - 2)
                        strTmp = Mid(strTmp, lngPos + 1)
                    End If
                    If .TextMatrix(LngRow, DI_��ϱ���) = "" And Not (IsNull(rsInput!���ID) And IsNull(rsInput!����id)) Then
                        '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                        .TextMatrix(LngRow, DI_��ϱ���) = IIf(Not IsNull(rsInput!����id), rsInput!����id & "", rsInput!���ID & "")
                    End If
                    '��ȡ��ҽ֤����������������ܻ�����ǰ��׺��ǰ��׺�������ţ����Է����ȡ�ַ���
                    If strTmp Like "?*(?*)" And Not bln��ҽ Then
                        strTmp = StrReverse(strTmp)
                        lngPos = InStr(1, strTmp, "(")
                        .TextMatrix(LngRow, DI_��ҽ֤��) = StrReverse(Mid(strTmp, 2, lngPos - 2))
                        strTmp = StrReverse(Mid(strTmp, lngPos + 1))
                    End If
                    
                    'ȡ�������
                    .TextMatrix(LngRow, DI_�������) = strTmp
                    '��������ı�������
                    If Not (IsNull(rsInput!���ID) And IsNull(rsInput!����id)) Then
                        .Cell(flexcpData, LngRow, DI_�������) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!������� & "")
                    Else
                        .Cell(flexcpData, LngRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                    End If
                    '���������ݼ���
                    .TextMatrix(LngRow, DI_�Ƿ�����) = IIf(Val(rsInput!�Ƿ����� & "") = 1, "��", "")
                    .TextMatrix(LngRow, DI_���ID) = rsInput!���ID & ""
                    .TextMatrix(LngRow, DI_����ID) = rsInput!����id & ""
                    .TextMatrix(LngRow, DI_֤��ID) = rsInput!֤��ID & ""
                    '.TextMatrix(LngRow, DI_ICD����) = rsInput!���� & ""
                    .TextMatrix(LngRow, DI_ҽ��IDs) = rsInput!ҽ��ID & ""
                    .TextMatrix(LngRow, DI_�����Դ) = Val(rsInput!��¼��Դ & "") '�����¼��Դ���Ա㱣��ʱ������Ϊ��ҳ�򲡰���Դ
                    If gclsPros.MedPageSandard <> ST_������ҳ Then
                        .TextMatrix(LngRow, DI_��ע) = rsInput!��ע & ""
                        .TextMatrix(LngRow, DI_��Ժ���) = rsInput!��Ժ��� & ""
                        .TextMatrix(LngRow, DI_��Ժ����) = rsInput!��Ժ���� & ""
                        .TextMatrix(LngRow, DI_�Ƿ�δ��) = IIf(Val(rsInput!�Ƿ�δ�� & "") = 1, "��", "")
                    Else
                        .TextMatrix(LngRow, DI_����ʱ��) = Format(rsInput!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                    End If
                    .RowData(LngRow) = Val(rsInput!ID & "")
                Else
                    .TextMatrix(LngRow, DI_����ID) = rsInput!����id & ""
                    .TextMatrix(LngRow, DI_ICD����) = rsInput!�������� & ""
                    .Cell(flexcpData, LngRow, DI_ICD����) = .TextMatrix(LngRow, DI_ICD����)
                End If
                rsInput.MoveNext
            Loop
        Next
        .Row = .FixedRows: .Col = DI_�������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'
Public Sub ArchiveLoadAller(ByVal vsAller As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'���ܣ����Ӳ������ļ��ع���ҩ��
    Dim i As Long, LngRow As Long

    rsInput.Filter = "��¼��Դ=3" '��ҳ������д��
    If rsInput.EOF Then rsInput.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
    With vsAller
        .Rows = rsInput.RecordCount + 1 '�̶���+����
        For i = 1 To rsInput.RecordCount
            '������Դ�Ŀ������ظ�
            LngRow = -1
            If Not IsNull(rsInput!ҩ��ID) Then
                LngRow = .FindRow(CLng(rsInput!ҩ��ID))
            ElseIf Not IsNull(rsInput!ҩ����) Then
                LngRow = .FindRow(CStr(rsInput!ҩ����), , AI_����ҩ��)
            End If
            If LngRow = -1 Then
                .RowData(i) = CLng(NVL(rsInput!ҩ��ID, 0))
                .TextMatrix(i, AI_����ʱ��) = Format(rsInput!����ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, AI_����ҩ��) = NVL(rsInput!ҩ����)
                .TextMatrix(i, AI_������Ӧ) = NVL(rsInput!������Ӧ)
            End If
            rsInput.MoveNext
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = AI_����ҩ��
    End With
End Sub

Public Sub ArchiveLoadOPS(ByVal vsOPS As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'���ܣ����Ӳ������ļ����������
    Dim i As Long

    With vsOPS
        .Rows = .FixedRows: .Rows = .FixedRows + rsInput.RecordCount + 1
        For i = 1 To rsInput.RecordCount
            .TextMatrix(i, PI_��������) = Format(rsInput!������ʼʱ�� & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_��������) = Format(rsInput!��������ʱ�� & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_��������) = rsInput!�������� & ""
            .TextMatrix(i, PI_��������) = rsInput!�������� & ""
            .TextMatrix(i, PI_����ҽʦ) = rsInput!����ҽʦ & ""
            .TextMatrix(i, PI_������ʿ) = rsInput!������ʿ & ""
            .TextMatrix(i, PI_����1) = rsInput!��һ���� & ""
            .TextMatrix(i, PI_����2) = rsInput!�ڶ����� & ""
            .TextMatrix(i, PI_����ʽ) = rsInput!����ʽ & ""
            .TextMatrix(i, PI_����ҽʦ) = rsInput!����ҽʦ & ""
            If rsInput!�п� & rsInput!���� & "" <> "" Then
                .TextMatrix(i, PI_�п�����) = rsInput!�п� & "/" & rsInput!����
            End If
            .TextMatrix(i, PI_��������ID) = Val(rsInput!��������ID & "")
            .TextMatrix(i, PI_������ĿID) = Val(rsInput!������Ŀid & "")
            .TextMatrix(i, PI_����ID) = Val(rsInput!����ID & "")
            .TextMatrix(i, PI_��������) = rsInput!�������� & ""
            .TextMatrix(i, PI_�������) = rsInput!������� & ""
            .TextMatrix(i, PI_ASA�ּ�) = rsInput!asa�ּ� & ""
            .TextMatrix(i, PI_NNIS�ּ�) = rsInput!NNIS�ּ� & ""
            .TextMatrix(i, PI_��������) = rsInput!�������� & ""
            .TextMatrix(i, PI_�ٴ�����) = IIf(Val(rsInput!�ٴ����� & "") = 1, -1, 0)
            .TextMatrix(i, PI_׼������) = Val(rsInput!׼������ & "")
            .TextMatrix(i, PI_������ҩʱ��) = Format(rsInput!������ҩʱ�� & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_����ʼʱ��) = Format(rsInput!����ʼʱ�� & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_�пڲ�λ) = rsInput!�пڲ�λ & ""
            .TextMatrix(i, PI_�ط�������Ŀ��) = rsInput!�ط�Ŀ�� & ""
            .TextMatrix(i, PI_�ط������Ҽƻ�) = IIf(Val(rsInput!�ط��ƻ� & "") = 1, -1, 0)
            .TextMatrix(i, PI_�пڸ�Ⱦ) = IIf(Val(rsInput!�пڸ�Ⱦ & "") = 1, -1, 0)
            .TextMatrix(i, PI_����֢) = IIf(Val(rsInput!����֢ & "") = 1, -1, 0)
            .Cell(flexcpData, i, PI_��������) = rsInput!����ԭ�� & ""
            .RowData(i) = Val(rsInput!ID & "")
            rsInput.MoveNext
        Next
    End With
End Sub

Public Sub ArchiveLoadKSS(ByVal vsKSS As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'���ܣ����Ӳ������ļ����������
    Dim LngRow As Long, i As Long

    With vsKSS
        Do While Not rsInput.EOF
           For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, KI_����ҩ����) = "" Or .RowData(i) = Val(rsInput!ҩ��id & "") Then
                    LngRow = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then
                .AddItem "": LngRow = i
            End If
            'װ������
            .RowData(LngRow) = Val(rsInput!ҩ��id & "")
            If .RowData(LngRow) <> 0 Then
                .TextMatrix(LngRow, KI_����ҩ����) = rsInput!���� & ""
                .Cell(flexcpData, LngRow, KI_����ҩ����) = .TextMatrix(LngRow, KI_����ҩ����)
                .TextMatrix(LngRow, KI_��ҩĿ��) = rsInput!��ҩĿ�� & ""
                .TextMatrix(LngRow, KI_ʹ�ý׶�) = rsInput!ʹ�ý׶� & ""
                .TextMatrix(LngRow, KI_ʹ������) = IIf(Val(rsInput!ʹ������ & "") = 0, "", Val(rsInput!ʹ������ & ""))
                .Cell(flexcpChecked, LngRow, KI_һ���п�Ԥ����) = Val(rsInput!һ���п�Ԥ���� & "")
                .TextMatrix(LngRow, KI_DDD��) = FormatEx(Val(rsInput!DDD�� & ""), 2)
                .TextMatrix(LngRow, KI_������ҩ) = rsInput!������ҩ & ""
            End If
            rsInput.MoveNext
        Loop
    End With
End Sub

Public Function ArchiveLoadPageMedRec(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ط����뻯����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 15:55:27
    '����:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    Set rsTmp = GetChemothData(lng����ID, lng��ҳID)
    Set vsTmp = gclsPros.CurrentForm.vsChemoth
    With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!��� & "")
                .TextMatrix(i, CI_��ѧ���Ʊ���) = NVL(rsTmp!������Ϣ)
                .TextMatrix(i, CI_��ʼ����) = Format(rsTmp!��ʼ����, "yyyy-MM-dd")
                .TextMatrix(i, CI_��������) = Format(rsTmp!��������, "yyyy-MM-dd")
                .TextMatrix(i, CI_�Ƴ���) = Format(Val(rsTmp!�Ƴ��� & ""), "###;-###;;")
                .TextMatrix(i, CI_����) = Format(Val(rsTmp!���� & ""), "###;-###;;")
                .TextMatrix(i, CI_���Ʒ���) = rsTmp!���Ʒ��� & ""
                .TextMatrix(i, CI_����Ч��) = rsTmp!����Ч�� & ""
                .TextMatrix(i, CI_����ID) = rsTmp!����id & ""
                rsTmp.MoveNext
            Next
    End With
    Set rsTmp = GetRadiothData(lng����ID, lng��ҳID)
    Set vsTmp = gclsPros.CurrentForm.vsRadioth
    With vsTmp
        .Rows = rsTmp.RecordCount + .FixedRows
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = Val(rsTmp!��� & "")
            .TextMatrix(i, RI_�������Ʊ���) = NVL(rsTmp!������Ϣ)
            .TextMatrix(i, RI_��ʼ����) = Format(rsTmp!��ʼ����, "yyyy-MM-dd")
            .TextMatrix(i, RI_��������) = Format(rsTmp!��������, "yyyy-MM-dd")
            .TextMatrix(i, RI_�������) = Format(Val(rsTmp!������� & ""), "###;-###;;")
            .TextMatrix(i, RI_�ۼ���) = Format(Val(rsTmp!�ۼ��� & ""), "###;-###;;")
            .TextMatrix(i, RI_��Ұ��λ) = rsTmp!��Ұ��λ & ""
            .TextMatrix(i, RI_����Ч��) = rsTmp!����Ч�� & ""
            .TextMatrix(i, RI_����ID) = rsTmp!����id & ""
            rsTmp.MoveNext
        Next
    End With
    If gclsPros.MedPageSandard = ST_��������׼ Then
        Set rsTmp = GetSpiritData(lng����ID, lng��ҳID)
        Set vsTmp = gclsPros.CurrentForm.vsSpirit
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!��� & "")
                .TextMatrix(i, SI_ҩ������) = rsTmp!ҩ������ & ""
                .TextMatrix(i, SI_�Ƴ�) = rsTmp!�Ƴ� & ""
                .TextMatrix(i, SI_�������) = rsTmp!������� & ""
                .TextMatrix(i, SI_���ⷴӦ) = rsTmp!���ⷴӦ & ""
                .TextMatrix(i, SI_��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, SI_ҩƷID) = rsTmp!ҩƷID & ""
                rsTmp.MoveNext
            Next
        End With
    End If
    ArchiveLoadPageMedRec = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ArchiveLoadOtherInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:���ظ�ҳ����
    '����:lng����id-����id
    '     lng��ҳid -��ҳid
    '����:���سɹ�,����true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim vsTmp As VSFlexGrid
    Err = 0: On Error GoTo Errhand
    '��֢�໤���
    If gclsPros.MedPageSandard <> ST_����ʡ��׼ Then
        Set rsTmp = GetICUData(lng����ID, lng��ҳID)
        If gclsPros.MedPageSandard <> ST_����ʡ��׼ Then
            Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
            With vsTmp
                .Rows = rsTmp.RecordCount + .FixedRows
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, UI_�໤������) = rsTmp!�໤������ & ""
                    .TextMatrix(i, UI_����ʱ��) = rsTmp!����ʱ�� & ""
                    .TextMatrix(i, UI_�˳�ʱ��) = rsTmp!�˳�ʱ�� & ""
                    If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                        .TextMatrix(i, UI_���) = Val(rsTmp!��� & "")
                        .Cell(flexcpChecked, i, UI_����ס�ƻ�) = Val(rsTmp!����ס�ƻ� & "")
                         .TextMatrix(i, UI_����סԭ��) = rsTmp!����סԭ�� & ""
                    End If
                    .RowData(i) = Val(rsTmp!��� & "")
                    rsTmp.MoveNext
                Next
            End With
        Else
            '���ϰ棬û�б��
            rsTmp.Sort = "���"
            If Not rsTmp.EOF Then
                For i = 0 To rsTmp.Fields.Count - 1
                    Call ArchiveSetCtrlValues(rsTmp.Fields(i).Name, rsTmp.Fields(i).Value & "")
                Next
            End If
        End If
    End If
    '��е���������ҽԺ��Ⱦ���걾�ͼ�
    If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
        Set rsTmp = GetICUInstrumentsData(lng����ID, lng��ҳID)
        Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, TI_ICU����) = rsTmp!�໤������ & ""
                .TextMatrix(i, TI_��е������) = rsTmp!��е������ & ""
                .TextMatrix(i, TI_��ʼʱ��) = rsTmp!��ʼʹ��ʱ�� & ""
                .TextMatrix(i, TI_����ʱ��) = rsTmp!����ʹ��ʱ�� & ""
                .TextMatrix(i, TI_��Ⱦ�ۼ�Сʱ) = rsTmp!��Ⱦ�ۼ�ʱ�� & ""
                .RowData(i) = Val(rsTmp!��� & "")
                rsTmp.MoveNext
            Next
        End With
        
        Set rsTmp = GetInfectData(lng����ID, lng��ҳID)
        Set vsTmp = gclsPros.CurrentForm.vsInfect
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, FI_ȷ������) = rsTmp!ȷ������ & ""
                .TextMatrix(i, FI_��Ⱦ��λ) = rsTmp!��Ⱦ��λ & ""
                .TextMatrix(i, FI_ҽԺ��Ⱦ����) = rsTmp!ҽԺ��Ⱦ���� & ""
                .RowData(i) = Val(rsTmp!��� & "")
                rsTmp.MoveNext
            Next
        End With
        
        Set rsTmp = GetSampleData(lng����ID, lng��ҳID)
        Set vsTmp = gclsPros.CurrentForm.vsSample
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, MI_�걾) = rsTmp!�걾 & ""
                .TextMatrix(i, MI_��ԭѧ���뼰����) = rsTmp!��ԭѧ���� & ""
                .TextMatrix(i, MI_�ͼ�����) = rsTmp!�ͼ����� & ""
                .RowData(i) = Val(rsTmp!��� & "")
                rsTmp.MoveNext
            Next
        End With
    End If
    ArchiveLoadOtherInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(gOldwinproc, hwnd, wMsg, wParam, lParam)
End Function


