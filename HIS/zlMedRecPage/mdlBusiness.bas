Attribute VB_Name = "mdlBusiness"
Option Explicit

Public Function InitTableDiag() As Boolean
'���ܣ�������ϱ����
    Dim strHeadXY As String, strHeadZY As String
    Dim strRowsXY As String, strRowsZY As String
    Dim intFixedRowsXY As Integer, intFixedRowsZY As Integer

    Dim intFixedColsXY As Integer, intFixedColsZY As Integer
    Dim vsTmp As VSFlexGrid
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    intFixedRowsXY = 1: intFixedColsXY = 1
    intFixedRowsZY = 1: intFixedColsZY = 1
    Select Case gclsPros.FuncType
        Case f���ѡ��
            If gclsPros.PatiType = PF_���� Then
                '��ʾ�У��������,����(���ѡ����),��ϱ���,�������,��ҽ֤��(��ҽ���),����ʱ��,����,����,ɾ��
                strHeadXY = "�������,900,4;����,450,4,11;��ϱ���,900,4;�������,4000,1;��ҽ֤��;����ʱ��,2200,1;��ע;��Ժ����;��Ժ���;ICD����;δ��;����,450,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strHeadZY = "�������,900,4;����,450,4,11;��ϱ���,900,4;�������,2900,1;��ҽ֤��,1500,1;����ʱ��,1800,1;��ע;��Ժ����;��Ժ���;ICD����;δ��;����,450,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strRowsXY = DI_������� & ",��ҽ," & DI_��Ϸ��� & "," & DT_�������XY
                strRowsZY = DI_������� & ",��ҽ," & DI_��Ϸ��� & "," & DT_�������ZY
            Else
                '��ʾ�У��������,����(���ѡ����),��ϱ���,�������,��ҽ֤��(��ҽ���),��ע,��Ժ����;��Ժ���,δ��,����,����,ɾ��
                strHeadXY = "����������ÿ�,1450,4;����,450,4,11;��ϱ���,850,4;�������,2500,1;��ҽ֤��;����ʱ��;��ע,1000,1;��Ժ����,850,1;��Ժ���,850,1;ICD����,700,1;δ��,450,4;����,450,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strHeadZY = "����������ÿ�,1450,4;����,450,4,11;��ϱ���,850,4;�������,1900,1;��ҽ֤��,1400,1;����ʱ��;��ע,900,1;��Ժ����,900,1;��Ժ���,900,1;ICD����;δ��;����,450,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strRowsXY = DI_������� & ",�ţ���������� ," & DI_��Ϸ��� & "," & DT_�������XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",Ժ�ڸ�Ⱦ," & DI_��Ϸ��� & "," & DT_Ժ�ڸ�Ⱦ & ";" & _
                                    DI_������� & ", �� �� ֢ ," & DI_��Ϸ��� & "," & DT_����֢ & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_������� & ";" & _
                                    DI_������� & ",�����ж�," & DI_��Ϸ��� & "," & DT_�����ж���
                strRowsZY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���ZY
            End If

        Case fҽ����ҳ
            '��ʾ�У��������,����(���ѡ����),��ϱ���,�������,��ҽ֤��(��ҽ���),��ע,��Ժ����;��Ժ���,δ��,����,����,ɾ��
            strHeadXY = "����������ÿ�,1350,4;����;��ϱ���,1000,4;�������,3200,1;��ҽ֤��;����ʱ��;��ע,1000,1;��Ժ����,850,1;��Ժ���,850,1;ICD����,800,1;δ��,450,4;����,450,4;" & _
                                    ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
            strHeadZY = "����������ÿ�,1350,4;����;��ϱ���,1000,4;�������,2700,1;��ҽ֤��,1500,1;����ʱ��;��ע,1300,1;��Ժ����,850,1;��Ժ���,850,1;ICD����;δ��;����,450,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
            strRowsXY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������XY & ";" & _
                                DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                DI_������� & ",Ժ�ڸ�Ⱦ," & DI_��Ϸ��� & "," & DT_Ժ�ڸ�Ⱦ & ";" & _
                                DI_������� & ", �� �� ֢ ," & DI_��Ϸ��� & "," & DT_����֢ & ";" & _
                                DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_������� & ";" & _
                                DI_������� & ",�����ж�," & DI_��Ϸ��� & "," & DT_�����ж���
            strRowsZY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������ZY & ";" & _
                                DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���ZY
        Case f������ҳ
                '��ʾ�У��������,��ϱ���,�������,��ҽ֤��(��ҽ���),��ע,��Ժ����;��Ժ���,ICD����,δ��,����,����,ɾ��
                strHeadXY = "����������ÿ�,1250,4;����;��ϱ���,900,4;�������,3200,1;��ҽ֤��;����ʱ��;��ע,1200,1;��Ժ����,850,1;��Ժ���,850,1;ICD����,800,1;δ��,350,4;����,350,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strHeadZY = "����������ÿ�,1250,4;����;��ϱ���,900,4;�������,3000,1;��ҽ֤��,1500,1;����ʱ��;��ע,1100,1;��Ժ����,850,1;��Ժ���,850,1;ICD����;δ��;����,350,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strRowsXY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",Ժ�ڸ�Ⱦ," & DI_��Ϸ��� & "," & DT_Ժ�ڸ�Ⱦ & ";" & _
                                    DI_������� & ", �� �� ֢ ," & DI_��Ϸ��� & "," & DT_����֢ & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_������� & ";" & _
                                    DI_������� & ",�����ж�," & DI_��Ϸ��� & "," & DT_�����ж���
                strRowsZY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���ZY
        Case f���Ӳ���
            If gclsPros.MedPageSandard = ST_������ҳ Then
                '��ʾ�У��������,��ϱ���,�������,��ҽ֤��(��ҽ���),����ʱ��,����
                strHeadXY = ",450,4;����;��ϱ���,900,4;�������,3000,1;��ҽ֤��;����ʱ��,1500,1;��ע;��Ժ����;��Ժ���;ICD����,800,1;δ��;����,450,4;" & _
                                        "����;ɾ��;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strHeadZY = ",450,4;����;��ϱ���,900,4;�������,3000,1;��ҽ֤��,1500,1;����ʱ��,1500,1;��ע;��Ժ����;��Ժ���;ICD����;δ��;����,450,4;" & _
                                        "����;ɾ��;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                intFixedRowsZY = 0
                strRowsXY = DI_������� & ",��ҽ," & DI_��Ϸ��� & "," & DT_�������XY
                strRowsZY = DI_������� & ",��ҽ," & DI_��Ϸ��� & "," & DT_�������ZY
            Else
                '��ʾ�У��������,��ϱ���,�������,��ҽ֤��(��ҽ���),��ע,��Ժ����;��Ժ���,δ��,����,����,ɾ��
                strHeadXY = "����������ÿ�,1350,4;����;��ϱ���,810,4;�������,2700,1;��ҽ֤��;����ʱ��;��ע,800,1;��Ժ����,1000,1;��Ժ���,810,1;ICD����,800,1;δ��,450,4;����,450,4;" & _
                                        "����;ɾ��;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strHeadZY = "����������ÿ�,1350,4;����;��ϱ���,810,4;�������,2500,1;��ҽ֤��,1050,1;����ʱ��;��ע,800,1;��Ժ����,1000,1;��Ժ���,810,1;ICD����;δ��,450,4;����,450,4;" & _
                                        "����;ɾ��;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
                strRowsXY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",Ժ�ڸ�Ⱦ," & DI_��Ϸ��� & "," & DT_Ժ�ڸ�Ⱦ & ";" & _
                                    DI_������� & ", �� �� ֢ ," & DI_��Ϸ��� & "," & DT_����֢ & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_������� & ";" & _
                                    DI_������� & ",�����ж�," & DI_��Ϸ��� & "," & DT_�����ж���
                strRowsZY = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���ZY
            End If
    End Select

    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    Call Grid.Init(vsTmp, strHeadXY, strRowsXY, intFixedColsXY, intFixedRowsXY)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            If Not .ColHidden(DI_��Ժ����) Then .ColData(DI_��Ժ���) = "��|�ٴ�δȷ��|�������|��"
            If Not .ColHidden(DI_��Ժ���) Then
                Set rsTmp = GetBaseCode("���ƽ��")
                If Not rsTmp.EOF Then
                    strTmp = Rec.ToComboList(rsTmp, "[0]-[1]|", "����", "����")
                    '��Chr(10)����հ�����Ϊ��ʵ�ַ��Ϳո񵯳������б�
                    .ColData(DI_��Ժ���) = Chr(10) & "|" & strTmp
                Else
                    .ColData(DI_��Ժ���) = Chr(10) & "|1-����|2-��ת|3-δ��|4-����|5-����"
                End If
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then
            .Font.Size = gclsPros.FontSize
            Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_���� & ",")
        End If
        If .TextMatrix(0, DI_�������) = "����������ÿ�" Then .TextMatrix(0, DI_�������) = "�������" '�ָ���ͷ
    End With

    Set vsTmp = gclsPros.CurrentForm.vsDiagZY
    Call Grid.Init(vsTmp, strHeadZY, strRowsZY, intFixedColsZY, intFixedRowsZY)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            If Not .ColHidden(DI_��Ժ����) Then .ColData(DI_��Ժ���) = "��|�ٴ�δȷ��|�������|��"
            If Not .ColHidden(DI_��Ժ���) Then
                If strTmp <> "" Then
                    '��Chr(10)����հ�����Ϊ��ʵ�ַ��Ϳո񵯳������б�
                    .ColData(DI_��Ժ���) = Chr(10) & "|" & strTmp
                Else
                    .ColData(DI_��Ժ���) = Chr(10) & "|1-����|2-��ת|3-δ��|4-����|5-����"
                End If
            End If
        End If
          If .Font.Size <> gclsPros.FontSize Then
             .Font.Size = gclsPros.FontSize
            Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_���� & ",")
          End If
        If .TextMatrix(0, DI_�������) = "����������ÿ�" Then .TextMatrix(0, DI_�������) = "�������" '�ָ���ͷ
    End With
    InitTableDiag = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableOPS() As Boolean
'���ܣ�����������������
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    Select Case gclsPros.MedPageSandard
        Case ST_��������׼
            strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "������ʼʱ��,1850,4;��������ʱ��,1850,4", "��������������,1850,4;��������ʱ��") & ";��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������������,1500,1;��������������,2800,1;�ٴ�����,850,4,11;����,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                            "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;��������,850,1;����ҽʦ,850,1;�п����ϵȼ�,1400,1;�пڲ�λ,850,1;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                            "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                            "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
        Case ST_����ʡ��׼
            strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "������ʼʱ��,1850,4;��������ʱ��,1850,4", "��������������,1850,4;��������ʱ��") & ";��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������������,1500,1;��������������,2800,1;�ٴ�����,850,4,11;����,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                            "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;��������,850,1;����ҽʦ,850,1;�п����ϵȼ�,1400,1;�пڲ�λ;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                            "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                            "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
        Case ST_�Ĵ�ʡ��׼
            strHead = ",300,4;" & "��ʼ����,1850,4;��������,1850,4;��ǰԤ���Կ�����ҩʱ��,2150,4;�������,875,1;׼������,850,7;��������,1500,1;��������,2800,1;�ٴ�����,850,4,11;����ҽʦ,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                            "����ʼʱ��,1550,4;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;�����ּ�,850,1;����ҽʦ,850,1;�п�/����,1400,1;�пڲ�λ,850,1;�ط������Ҽƻ�,1400,4,11;�ط�������Ŀ��,1400,1;�пڸ�Ⱦ,850,4,11;����֢,720,4,11;" & _
                            "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                            "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
        Case ST_����ʡ��׼
            strHead = ",300,4;" & "��������,1850,4;��������;��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������,1500,1;��������,2800,1;�ٴ�����,850,4,11;����ҽʦ,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                            "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;�����ּ�,850,1;����ҽʦ,850,1;�п�/����,1400,1;�пڲ�λ;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                            "��ǰ0.5-2СʱԤ���ÿ���ҩ,2400,4,11;�������Χ����Ԥ���ÿ���ҩ����,2850,7;��Ԥ�ڵĶ�������,1600,4,11;������֢,1000,4,11;������������,1200,4,11;��������֢,1000,4,11;" & _
                            "�����Ѫ��Ѫ��,1450,4,11;�����˿��ѿ�,1200,4,11;�������Ѫ˨,1450,4,11;��������/��л����,1700,4,11;�������˥��,1200,4,11;�����˨��,1000,4,11;�����Ѫ֢,1000,4,11;" & _
                            "�����Źؽڹ���,1450,4,11;��������ID;������ĿID;����ID;��������;������Դ"
    End Select
    Set vsTmp = gclsPros.CurrentForm.vsOPS
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        .Font.Size = 9
        If gclsPros.FuncType <> f���Ӳ��� Then
            .ColComboList(PI_�������) = " |����|����|����"
            .ColComboList(PI_ASA�ּ�) = " |P1|P2|P3|P4|P5|P6"
            .ColComboList(PI_NNIS�ּ�) = " |NNIS0��|NNIS1��|NNIS2��|NNIS3��"
            .ColComboList(PI_��������) = " |��|һ������|��������|��������|�ļ�����"
            '�п�����
            Set rsTmp = GetBaseCode("�����п�����")
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "����", "����")
            Else
                strTmp = " |0-0 / |1-��/��|2-��/��|3-��/��|4-��/����|5-��/��|6-��/��|7-��/��|8-��/����|9-��/��|10-��/��|11-��/��|12-��/����|13-IV/��|14-IV/��|15-IV/��|16-IV/����"
            End If
            .ColData(PI_�п�����) = strTmp
            '��������
            Set rsTmp = GetBaseCode("������������")
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "����", "����")
            Else
                strTmp = " |JM-����|QM-ȫ��|CY-��Ӳ|QT-����|JM-����|BC-�۴�|JC-����"
            End If
            .ColData(PI_��������) = strTmp
        End If
        If gclsPros.FontSize <> 9 Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableOPS = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableAller() As Boolean
'���ܣ����ù�����������
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.FuncType = f���Ӳ��� Then
        strHead = "������,4500,1;������Ӧ,2000,1;����ʱ��,1500,4;����Դ����;ҩ��ID;������Դ "
    Else
        strHead = "������,4500,1;������Ӧ,4500,1;����ʱ��,1500,4;����Դ����;ҩ��ID;������Դ "
    End If
    Set vsTmp = gclsPros.CurrentForm.vsAller
    Call Grid.Init(gclsPros.CurrentForm.vsAller, strHead)

    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableAller = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableKSS() As Boolean
'���ܣ����ÿ�����ʹ����������
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.FuncType = f���Ӳ��� Then
        strHead = ",300,4;����ҩ������,2250,1;��ҩĿ��,2000,1;ʹ�ý׶�,1100,1;ʹ������,850,7;�����п�Ԥ����,1400,4,11;DDD��,800,7;������ҩ,900,1"
    Else
        strHead = ",300,4;����ҩ������,3000,1;��ҩĿ��,950,1;ʹ�ý׶�,900,1;ʹ������,950,7;�����п�Ԥ����,1400,4,11;DDD��,1000,7;������ҩ,900,1"
    End If
    Set vsTmp = gclsPros.CurrentForm.vsKSS
    Call Grid.Init(vsTmp, strHead, , 1)
    With vsTmp
        .Font.Size = 9
        If gclsPros.FuncType <> f���Ӳ��� Then
            .ColComboList(KI_����ҩ����) = "..."
            .ColComboList(KI_ʹ�ý׶�) = " |��ǰ|����|����|Χ������"
            .ColComboList(KI_������ҩ) = "����|����|����|����|>����"
            .ColComboList(KI_��ҩĿ��) = " |Ԥ��|����"
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableKSS = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTablefMain() As Boolean
'���ܣ����ò���������Ŀ�����
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim LngCol As Long, LngRow As Long, i As Long, lngCount As Long

    On Error GoTo errH
    If gclsPros.FuncType = f���Ӳ��� Then
         '�Ĵ���3����ʾ�������汾2����ʾ
        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
            strHead = "��Ŀ,1500,4;����,1800,1;ֵ��;��Ŀ,1500,4;����,1800,1;ֵ��;��Ŀ,1500,4;����,1800,1;ֵ��"
        Else
            strHead = "��Ŀ,1500,4;����,1250,1;ֵ��;��Ŀ,1500,4;����,1250,1;ֵ��"
        End If
    ElseIf gclsPros.FuncType = f������ҳ Then
        strHead = "��Ŀ,1620,4;����,2210,1;ֵ��;��Ŀ,1620,4;����,2210,1;ֵ��;��Ŀ,1620,4;����,2210,1;ֵ��"
    ElseIf gclsPros.FuncType = fҽ����ҳ Then
        strHead = "��Ŀ,1600,4;����,2030,1;ֵ��;��Ŀ,1600,4;����,2030,1;ֵ��;��Ŀ,1600,4;����,2030,1;ֵ��"
    End If

    
    Set vsTmp = gclsPros.CurrentForm.vsfMain
    Call Grid.Init(vsTmp, strHead)
    strSql = "Select Rownum ���, ����, ���� From (Select ����, ���� From ������Ŀ Order By ����)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption)
    With vsTmp
        If rsTmp.RecordCount = 0 Then
            .Rows = .FixedRows
        Else
            If gclsPros.FuncType = f���Ӳ��� Then
                lngCount = IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, 3, 2)
            Else
                lngCount = 3
            End If
            .Rows = rsTmp.RecordCount \ lngCount + 1 + IIf(rsTmp.RecordCount Mod lngCount = 0, 0, 1)
            For i = 0 To .Cols - 1 Step 3
                .Cell(flexcpBackColor, 1, i, .Rows - 1, i) = &HFCE7D8
                .FixedAlignment(i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = flexAlignLeftCenter
            Next
            Do While Not rsTmp.EOF
                i = Val(rsTmp!��� & "")
                LngRow = .FixedRows + ((i - 1) \ lngCount): LngCol = ((i - 1) Mod lngCount) * 3
                .TextMatrix(LngRow, LngCol) = rsTmp!����
                .TextMatrix(LngRow, LngCol + 2) = rsTmp!���� & ""
                If rsTmp!���� & "" = "�Ƿ�" Then
                    .TextMatrix(LngRow, LngCol + 1) = "��"
                    .Cell(flexcpChecked, LngRow, LngCol + 1) = 2
                    .Cell(flexcpAlignment, LngRow, LngCol + 1) = flexAlignCenterCenter
                Else
                    .Cell(flexcpAlignment, LngRow, LngCol + 1) = flexAlignLeftCenter
                End If
                rsTmp.MoveNext
            Loop
            If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
        End If
    End With
    InitTablefMain = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableChemoth() As Boolean
'���ܣ����û�����Ŀ�����
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If Not gclsPros.ReadPages Then InitTableChemoth = True: Exit Function
    If gclsPros.FuncType = f���Ӳ��� Then
        strHead = "��ѧ���Ʊ���,2750,1;��ʼ����,1000,4;��������,1000,4;�Ƴ���,900,7;���Ʒ���,2000,1;����,900,7;����Ч��" & vbNewLine & "(CR PR NC PD),900,4;����ID"
    ElseIf gclsPros.FuncType = f������ҳ Then
        strHead = "��ѧ���Ʊ���,3400,1;��ʼ����,1400,4;��������,1400,4;�Ƴ���,700,7;���Ʒ���,2500,1;����,800,7;����Ч��" & vbNewLine & "(CR PR NC PD),500,4;����ID"
    ElseIf gclsPros.FuncType = fҽ����ҳ Then
        strHead = "��ѧ���Ʊ���,3400,1;��ʼ����,1200,4;��������,1200,4;�Ƴ���,700,7;���Ʒ���,2500,1;����,800,7;����Ч��" & vbNewLine & "(CR PR NC PD),500,4;����ID"
    End If

    Set vsTmp = gclsPros.CurrentForm.vsChemoth
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            .ColComboList(CI_����Ч��) = "CR|PR|NC|PD"
            strTmp = zlDatabase.GetPara("������Ŀ", 300, 200, "")
            If strTmp <> "" Then
                '˵��:������Ŀ��Ϣ���Լ�������Ϊ׼,��ʽΪ:��������,ȱʡ��־;��������1,ȱʡ��־1;...
                strSql = "Select /*+ Rule*/" & vbNewLine & _
                        " a.Id, a.����, a.���� || '-' || a.���� As ������Ϣ, b.ȱʡ��־,a.���" & vbNewLine & _
                        "From ��������Ŀ¼ A," & vbNewLine & _
                        "     (Select C1 ����, (Case Instr(C2, ',') When 0 Then C2 Else Substr(C2, 1, Instr(C2, ',') - 1) end) As ȱʡ��־," & vbNewLine & _
                        "             (Case Instr(C2, ',') When 0 Then '1' Else Substr(C2, Instr(C2, ',') +1) end) As ���" & vbNewLine & _
                        "       From Table(f_Str2list2([1], ';', ','))) B" & vbNewLine & _
                        "Where a.���� = b.���� And A.���=B.��� And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ŀ��Ϣ", strTmp)
                    .ColComboList(CI_��ѧ���Ʊ���) = .BuildComboList(rsTmp, "������Ϣ", "ID")
                    If rsTmp.RecordCount = 1 Then
                        .ColData(CI_��ѧ���Ʊ���) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!������Ϣ)
                    ElseIf rsTmp.RecordCount > 1 Then
                        rsTmp.Filter = "ȱʡ��־ like '1*'"
                        If rsTmp.EOF = False Then
                            .ColData(CI_��ѧ���Ʊ���) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!������Ϣ)
                        End If
                    Else
                        .ColData(CI_��ѧ���Ʊ���) = ";"
                        gclsPros.CurrentForm.lblEdit(0).Caption = "û�п��õĻ������Ʊ��룬�뵽����ϵͳ�����á�"
                        gclsPros.CurrentForm.lblEdit(0).Visible = True
                        .Editable = flexEDNone
                    End If
            Else
                .ColData(CI_��ѧ���Ʊ���) = ";"
                gclsPros.CurrentForm.lblEdit(0).Caption = "û�п��õĻ������Ʊ��룬�뵽����ϵͳ�����á�"
                gclsPros.CurrentForm.lblEdit(0).Visible = True
                .Editable = flexEDNone
            End If
        End If
    End With
    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableChemoth = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableRadioth() As Boolean
'���ܣ����÷�����Ŀ�����
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If Not gclsPros.ReadPages Then InitTableRadioth = True: Exit Function
    If gclsPros.FuncType = f���Ӳ��� Then
        strHead = "�������Ʊ���,2750,1;��ʼ����,1000,4;��������,1000,4;��Ұ��λ,2300,1;�������,900,7;�ۼ���,1000,7;����Ч��,900,4;����ID"
    ElseIf gclsPros.FuncType = f������ҳ Then
        strHead = "�������Ʊ���,3400,1;��ʼ����,1400,4;��������,1400,4;��Ұ��λ,2300,1;�������,900,7;�ۼ���,1000,7;����Ч��,600,4;����ID"
    ElseIf gclsPros.FuncType = fҽ����ҳ Then
        strHead = "�������Ʊ���,3400,1;��ʼ����,1200,4;��������,1200,4;��Ұ��λ,2300,1;�������,900,7;�ۼ���,1000,7;����Ч��,600,4;����ID"
    End If
    
    Set vsTmp = gclsPros.CurrentForm.vsRadioth
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            .ColComboList(RI_����Ч��) = "CR|PR|NC|PD"
            strTmp = zlDatabase.GetPara("������Ŀ", 300, 200, "")
            If strTmp <> "" Then
                '˵��:������Ŀ��Ϣ���Լ�������Ϊ׼,��ʽΪ:��������,ȱʡ��־;��������1,ȱʡ��־1;...
                strSql = "Select /*+ Rule*/" & vbNewLine & _
                        " a.Id, a.����, a.���� || '-' || a.���� As ������Ϣ, b.ȱʡ��־,a.���" & vbNewLine & _
                        "From ��������Ŀ¼ A," & vbNewLine & _
                        "     (Select C1 ����, (Case Instr(C2, ',') When 0 Then C2 Else Substr(C2, 1, Instr(C2, ',') - 1) end) As ȱʡ��־," & vbNewLine & _
                        "             (Case Instr(C2, ',') When 0 Then '1' Else Substr(C2, Instr(C2, ',') +1) end) As ���" & vbNewLine & _
                        "       From Table(f_Str2list2([1], ';', ','))) B" & vbNewLine & _
                        "Where a.���� = b.���� And A.���=B.��� And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ŀ��Ϣ", strTmp)
                    .ColComboList(RI_�������Ʊ���) = .BuildComboList(rsTmp, "������Ϣ", "ID")
                    If rsTmp.RecordCount = 1 Then
                        .ColData(RI_�������Ʊ���) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!������Ϣ)
                    ElseIf rsTmp.RecordCount > 1 Then
                        rsTmp.Filter = "ȱʡ��־ like '1*'"
                        If rsTmp.EOF = False Then
                            .ColData(RI_�������Ʊ���) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!������Ϣ)
                        End If
                    Else
                        .ColData(RI_�������Ʊ���) = ";"
                        gclsPros.CurrentForm.lblEdit(1).Caption = "û�п��õķ������Ʊ��룬�뵽����ϵͳ�����á�"
                        gclsPros.CurrentForm.lblEdit(1).Visible = True
                        .Editable = flexEDNone
                    End If
            Else
                .ColData(RI_�������Ʊ���) = ";"
                gclsPros.CurrentForm.lblEdit(1).Caption = "û�п��õķ������Ʊ��룬�뵽����ϵͳ�����á�"
                gclsPros.CurrentForm.lblEdit(1).Visible = True
                .Editable = flexEDNone
            End If
        End If
    End With
    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableRadioth = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableFlxAddICU() As Boolean
'���ܣ�����ICU��ס��������
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard = ST_��������׼ Then
        If gclsPros.FuncType = fҽ����ҳ Then
            strHead = "���;��֢�໤������,7000,4;����ʱ��(_��_��_�� ʱ_��_),2000,1,4,9999-99-99 99:99;�˳�ʱ��(_��_��_�� ʱ_��_),2000,1,4,9999-99-99 99:99;����ס�ƻ�;����סԭ��"
        ElseIf gclsPros.FuncType = f������ҳ Then
            strHead = "���;��֢�໤������,6500,4;����ʱ��(_��_��_�� ʱ_��_),2500,1,4,9999-99-99 99:99;�˳�ʱ��(_��_��_�� ʱ_��_),2500,1,4,9999-99-99 99:99;����ס�ƻ�;����סԭ��"
        Else
            strHead = "���;��֢�໤������,3000,4;����ʱ��(_��_��_�� ʱ_��_),2800,1,4,9999-99-99 99:99;�˳�ʱ��(_��_��_�� ʱ_��_),2800,1,4,9999-99-99 99:99;����ס�ƻ�;����סԭ��"
        End If
    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
        strHead = "���,450,7;ICU����,3100,1;��סʱ��,2100,4,7,9999-99-99 99:99;ת��ʱ��,2100,4,7,9999-99-99 99:99;����ס�ƻ�,1200,4,11;����סԭ��,800,1"
    Else
        InitTableFlxAddICU = True: Exit Function
    End If
    Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
    Call Grid.Init(vsTmp, strHead, , IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, 1, 0))
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            If gclsPros.MedPageSandard = ST_��������׼ Then
                .ColComboList(UI_�໤������) = "..."
            Else
                .ColComboList(UI_�໤������) = Rec.ToComboList(GetBaseCode("ICU����"), "[0].[1]|", "����", "����")
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableFlxAddICU = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableSpirit() As Boolean
'���ܣ����þ���ҩƷʹ����������
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_��������׼ Then InitTableSpirit = True: Exit Function
    If Not gclsPros.ReadPages Then InitTableSpirit = True: Exit Function
    strHead = "ҩ������,2500,1;�Ƴ�,2000,1;�������,1500,7;���ⷴӦ,2000,1;��Ч,2000,1;ҩƷid"
    Set vsTmp = gclsPros.CurrentForm.vsSpirit
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableSpirit = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableTSJC() As Boolean
'���ܣ����þ���ҩƷʹ����������
    Dim strHead As String
    Dim strRows As String
    Dim intFixedRows As Integer, intFixedCols As Integer
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    
    If gclsPros.FuncType = f���Ӳ��� Then
        strHead = ",1000,1;,2600,1"
    Else
        strHead = ",1250,1;,2600,1"
    End If
     
    If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
        strRows = "0,CT;0,PETCT;0,˫ԴCT;0,XƬ;0,B��;0,�����Ķ�ͼ;0,MRI;0,ͬλ�ؼ��"
    Else
        strRows = "0,������4;0,������5;0,������6"
    End If
    Set vsTmp = gclsPros.CurrentForm.vsTSJC
    Call Grid.Init(vsTmp, strHead, strRows, 1, 0)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                .ColComboList(1) = "1-����|2-����|3-δ��"
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableTSJC = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableICUInstruments() As Boolean
'���ܣ����þ���ҩƷʹ����������
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then InitTableICUInstruments = True: Exit Function
    strHead = "ICU����,3100,1;��е�򵼹�����,2400,1;��ʼʹ��ʱ��,1600,4,7,9999-99-99 99:99;����ʹ��ʱ��,1600,4,7,9999-99-99 99:99;��Ⱦ�ۼ�ʱ��(Сʱ:����),1100,7,,9999:99;���"
    Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then .ColComboList(TI_��е������) = Rec.ToComboList(GetBaseCode("��е����Ŀ¼"), "[0].[1]|", "����", "����")
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableICUInstruments = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableInfect() As Boolean
'���ܣ�����ҽԺ��Ⱦ��������
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then InitTableInfect = True: Exit Function
    strHead = "ȷ������,1400,4,,9999-99-99;��Ⱦ��λ,1400,1;ҽԺ��Ⱦ����,1000,1;ҽԺ��Ⱦ����"
    Set vsTmp = gclsPros.CurrentForm.vsInfect
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then .ColComboList(FI_��Ⱦ��λ) = Rec.ToComboList(GetBaseCode("��Ⱦ��λ"), "[0].[1]|", "����", "����")
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableInfect = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableSample() As Boolean
'���ܣ����ñ걾��Դ�����
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then InitTableSample = True: Exit Function
    strHead = "�걾,1400,1;��ԭѧ���뼰����,2800,1;�ͼ�����,1200,4,7,9999-99-99"
    Set vsTmp = gclsPros.CurrentForm.vsSample
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f���Ӳ��� Then
            .ColComboList(MI_�걾) = "1.ѪҺ|2.��Һ|3.���|4.̵Һ|5.����������"
            .ColComboList(MI_��ԭѧ���뼰����) = Rec.ToComboList(GetBaseCode("��ԭѧĿ¼"), "[0]-[1]|", "����", "����")
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableSample = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableFees() As Boolean
'���ܣ����÷���ͳ�Ʊ����
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If gclsPros.FuncType <> f������ҳ Then InitTableFees = True: Exit Function
    strHead = "������,2820,1;���ý��,1000,7;������,2820,1;���ý��,1000,7;������,2820,1;���ý��,1000,7"
    Set vsTmp = gclsPros.CurrentForm.vsFees
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.OpenMode <> EM_���� Then
            '��ѯ��ʽ�²���Ҫ���г�ʼ��
            '57638:������,2013-05-02,������Ҫ��ʾ�¼�
            strSql = "Select �ϼ� || Decode(NVL(�ϼ�,''),'','', '_') || ���� ����,���� From ������Ŀ  START WITH �ϼ� IS NULL CONNECT BY PRIOR ���� = �ϼ� ORDER BY �ϼ� || ����"
            Call zlDatabase.OpenRecordset(rsTmp, strSql, gclsPros.CurrentForm.Caption)
            strTmp = Rec.ToComboList(rsTmp, "[0].[1]|", "����", "����")
            If strTmp <> "" Then
                .ColComboList(0) = strTmp
                .ColComboList(2) = strTmp
                .ColComboList(4) = strTmp
            End If
        End If
    End With
    InitTableFees = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function LoadMedPageData(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long = 1, Optional ByVal str�Һŵ� As String, Optional blnNotReRead As Boolean = False, Optional ByVal bln��Ŀ As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------
    '����:����ҳ���ݼ��ص�������
    '���:lng����ID=����ID
    '     lng��ҳID=������ҳID
    '     blnNotReRead=�Ƿ��ǳ�ʼ���ݼ���,Fasle=���ǳ�ʼ���أ�True=�ǳ�ʼ����
    '     bln��Ŀ=�Ƿ��ȡ��Ŀ�����ݣ��Բ���ϵͳ��Ч
    '����:
    '����:���ط��ü�¼��
    '����:��˶
    '����:2013-12-26 10:43:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strCode As String
    Dim intMaxDiagSource As Integer
    Dim strTmp As String
    On Error GoTo errH
    Screen.MousePointer = 11
    With gclsPros.CurrentForm
        Set gclsPros.PatiInfo = GetPatiMainInfoData(lng����ID, lng��ҳID, IIf(gclsPros.PatiType = PF_����, gclsPros.RegistNo, "")) '������ҳ�Լ�������Ϣ
        '���ﲡ�˿��ܣ�������ҳ��ֻ���˹Һŵ��������Ҫ������������ID����ҳID����
        If gclsPros.PatiType = PF_���� Then
            lng����ID = gclsPros.����ID
            lng��ҳID = gclsPros.��ҳID
        End If
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(lng����ID, lng��ҳID, IIf(gclsPros.PatiType = PF_����, gclsPros.RegistNo, "")) '�ӱ���Ϣ
        If gclsPros.FuncType = f������ҳ Then
            '��ʼ��������Ϣ��¼��
            Set grsDeliceryInfo = zlDatabase.CopyNewRec(gclsPros.AuxiInfo, True, "��Ϣ��,��Ϣֵ,��Ϣֵ ��Ϣ��ֵ", Array("����", adInteger, 1, Empty, "��¼����", adInteger, 1, Empty))
            Set grsBabyDiag = zlDatabase.CopyNewRec(GetBabyDiagData(lng����ID, lng��ҳID), , , Array("��¼����", adInteger, 1, Empty))
            Set grsBabyInfo = zlDatabase.CopyNewRec(GetBabyInfoData(lng����ID, lng��ҳID), , , Array("��¼����", adInteger, 1, Empty))
        End If
        
        '���ز�����Ϣ
        If Not gclsPros.PatiInfo.EOF Then
            For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
                 Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
            Next
        End If
        
        '���ز��˴ӱ���Ϣ�Ͳ���������Ŀ
        If Not gclsPros.AuxiInfo.EOF Then
            gclsPros.AuxiInfo.MoveFirst
            For i = 1 To gclsPros.AuxiInfo.RecordCount
                Call SetCtrlValues(gclsPros.AuxiInfo!��Ϣ�� & "", gclsPros.AuxiInfo!��Ϣֵ & "", gclsPros.AuxiInfo!���� & "")
                gclsPros.AuxiInfo.MoveNext
            Next
        End If
        
        '104684�����޸ģ���Ժʱ��ȡֵ���ʱ��
        strTmp = GetInDeptTime(lng����ID, lng��ҳID, "____-__-__ __:__")
        If IsDate(strTmp) Then
            .mskDateInfo(DC_��Ժʱ��).Text = strTmp
            .txtDateInfo(DC_��Ժʱ��).Text = .mskDateInfo(DC_��Ժʱ��).Text
            gclsPros.InTime = Format(strTmp, "yyyy-MM-dd hh:mm:ss")
            strTmp = ""
        End If
        
        If gclsPros.FuncType = f������ҳ Then '������ҳ
            '������ҳסԺ�ţ������ţ������ŵȵ�����
            If gclsPros.OpenMode = EM_������ҳ Or gclsPros.OpenMode = EM_�������� Then
                strCode = gclsPros.PatiInfo!��Ժ���ұ��� & ""
                If strCode = "" Then strCode = gclsPros.PatiInfo!�����ұ��� & ""
                'סԺ�Ż�ȡ
                If IsNull(gclsPros.PatiInfo!סԺ��) Then
                    gclsPros.InNo = NVL(GetNextNo(2))
                ElseIf gclsPros.NewInNo And IsHavePageNos(CT_סԺ��, Not gclsPros.OpenMode = EM_�༭ Or gclsPros.Is��Ŀ, gclsPros.PatiInfo!סԺ�� & "", gclsPros.����ID) Then
                    gclsPros.InNo = NVL(GetNextNo(2))
                Else
                    gclsPros.InNo = gclsPros.PatiInfo!סԺ�� & ""
                End If
                .txtSpecificInfo(SLC_סԺ��).Text = gclsPros.InNo
                '�����Ż�ȡ
                If IsNull(gclsPros.PatiInfo!������) Then
                    If gclsPros.NewInNo Or Not gclsPros.SinPageNo And IsNull(gclsPros.PatiInfo!��󲡰���) Then
                        '�����ʹ���µ�סԺ��,������ǿ��Ĭ��ΪסԺ��
                        '���������סԺ������ , �򲡰��� = ��ǰסԺ��
                        .txtInfo(GC_������).Text = .txtSpecificInfo(SLC_סԺ��).Text
                    ElseIf gclsPros.SinPageNo Then
                        .txtInfo(GC_������).Text = NVL(GetNextNo(4, , strCode))
                    ElseIf Not IsNull(gclsPros.PatiInfo!��󲡰���) Then
                        '�����ǰ����������סԺ������,��ȡ���һ���������Ĳ�����
                        .txtInfo(GC_������).Text = gclsPros.PatiInfo!��󲡰��� & ""
                    End If
                Else
                    .txtInfo(GC_������).Text = gclsPros.PatiInfo!������ & ""
                End If
            End If
            If gclsPros.OpenMode <> EM_���� Then
                If IsNull(gclsPros.PatiInfo!��󵵰���) And gclsPros.UseFileRules Then
                    .txtInfo(GC_������).Text = NVL(GetNextNo(5, , strCode))
                Else
                    .txtInfo(GC_������).Text = gclsPros.PatiInfo!��󵵰��� & ""
                End If
            End If
            If gclsPros.Is��Ŀ Then
            '��ȡδ��Ŀ�����ݣ�ֻ�ܴ�ϵͳ��Ϣ�����ϳ�����Ҫ������
                '����26071 by lesfeng 2009-11-29 ����Ѫ�����ϵͳ��ȡ��Ѫ��Ϣ
                Call GetBloodValue(lng����ID, lng��ҳID)
                If gclsPros.OnLine Then
                    '����:ȡ������Ϣ:2009-02-03 14:51:23:14878
                    Call GetCareValue(lng����ID, lng��ҳID)
                    'סԺת����Ϣ
                    Call LoadTransferData(GetPatiTransfer(lng����ID, lng��ҳID))
                End If
            End If
            '���ط�����Ϣ
            Call CacheLoadVsFreesData(.vsFees, GetFreeData(lng����ID, lng��ҳID, Not gclsPros.Is��Ŀ), , Not gclsPros.Is��Ŀ)
        Else
             If gclsPros.PatiType = PF_���� Then '������ҳ��Ϣ����
                '����������Ϣ����
                Call .UCPatiVitalSigns.LoadPatiVitalSigns(lng����ID, lng��ҳID)
                '������Ƭ����
                Call ReadPatPricture(lng����ID, .imgPatient, strTmp)
                gclsPros.PictureFile = strTmp
                gclsPros.CurrentForm.picPatient.Tag = strTmp
             Else 'סԺ��ҳ
                '�ٴ�·�������Ϣ��ȡ
                 Call GetPatiPathInfo
                
                'סԺת����Ϣ
                Call LoadTransferData(GetPatiTransfer(lng����ID, lng��ҳID))
                '�Զ���ȡת�ƿ��Ҽ��������(���� ��)
                If .txtInfo(GC_��Ժ����).Text = "" Or .txtInfo(GC_��Ժ����).Text = "" Then
                    Set rsTmp = GetPatiRoom(lng����ID, lng��ҳID)
                    If .txtInfo(GC_��Ժ����).Text = "" Then .txtInfo(GC_��Ժ����).Text = rsTmp!��Ժ���� & ""
                    If .txtInfo(GC_��Ժ����).Text = "" Then .txtInfo(GC_��Ժ����).Text = rsTmp!��Ժ���� & ""
                End If
             End If
        End If
        '������ҳ��סԺ��ҳ������Ϣ����
        If gclsPros.PatiType <> PF_���� Then
            '����Ϣ�������������Ϣ�໥Ӱ����������Ϣ�Ĵ���
            '�����ȴ������гɹ�����,�Ѿ����أ�������Ҫ���
            If Val(gclsPros.PatiInfo!���ȴ��� & "") = 0 Then
                .txtSpecificInfo(SLC_���ȴ���).Text = ""
                .txtSpecificInfo(SLC_�ɹ�����).Text = ""
            End If
            '����ʱ��������������
            If Val(gclsPros.PatiInfo!�����־ & "") <> 0 Then
                .cboSpecificInfo(SLC_��������).Text = decode(Val(gclsPros.PatiInfo!�����־ & ""), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����", -1)
                .txtSpecificInfo(SLC_��������).Text = decode(Val(gclsPros.PatiInfo!�����־ & ""), 0, "", 9, "", NVL(gclsPros.PatiInfo!��������, 0))
                Call CboSpecificInfoClick(SLC_��������)
            End If
        End If
        '��סԺ����ID,��Ժ����ID,��ͬ��λID�����ڽ���ؼ���,����ʱ���ܻ��õ�
        .txtAdressInfo(ADRC_��λ��ַ).Tag = gclsPros.PatiInfo!��ͬ��λid & ""
        If gclsPros.PatiType = PF_סԺ Then
            .txtInfo(GC_��Ժ����).Tag = gclsPros.PatiInfo!��Ժ����ID & ""
            .txtInfo(GC_��Ժ����).Tag = gclsPros.PatiInfo!��Ժ����ID & ""
        End If
         '������Ϣ����
         If gclsPros.MedPageSandard <> ST_������ҳ Then
            If .chkInfo(CHK_�޹�����¼).Value = 0 Then
                Set rsTmp = GetAllerData(lng����ID, lng��ҳID)
                Call CacheLoadVsAllerData(.vsAller, rsTmp)
            End If
        ElseIf .chkInfo(CHK_�޹�����¼).Value = 0 Then '��ѡ�޹�����¼���򲻼��ع�����¼
            Set rsTmp = GetAllerData(lng����ID, lng��ҳID)
            Call CacheLoadVsAllerData(.vsAller, rsTmp)
        End If
        '��ȡ���
        Set rsTmp = GetPatiDiagData(lng����ID, lng��ҳID, IIf(gclsPros.PatiType <> PF_����, 1, 0), , Not gclsPros.Is��Ŀ, gclsPros.Moved)
        rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
        intMaxDiagSource = IIf(gclsPros.FuncType = f������ҳ, 4, -1)
        If gclsPros.FuncType = f������ҳ And rsTmp.EOF Then
            intMaxDiagSource = 3
            rsTmp.Filter = "��¼��Դ=3"
            If rsTmp.EOF Then intMaxDiagSource = 2
        End If
        If Not gclsPros.Is���� Or gclsPros.Is���� And rsTmp.RecordCount = 0 Then
            '����޸Ķ����Ժ���˲�������ҽ���ʱ��������ϴ��ҵ�����
            gclsPros.MainInfoRec.Filter = "��Ϣ��='��ҽ���' or ��Ϣ��='��ҽ���'"
            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
            If gclsPros.SecdInfoRec.RecordCount > 0 Then
                gclsPros.SecdInfoRec.MoveFirst
                For i = 1 To gclsPros.SecdInfoRec.RecordCount
                    gclsPros.SecdInfoRec.Delete
                    gclsPros.SecdInfoRec.MoveNext
                Next
            End If
            '2��������ҽ���
            '   1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
            Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "1,2,3,5,6,7,10", "1"), , intMaxDiagSource)
            '3��������ҽ���
            '   11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
            If gclsPros.Have��ҽ Then
                Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "11,12,13", "11"), , intMaxDiagSource)
            End If
        End If
        'סԺ��ҳ������ҳ������
        If gclsPros.PatiType <> PF_���� Then
            '���ز�ԭѧ���
            Call FilterDiagByType(rsTmp, DT_��ԭѧ���, intMaxDiagSource)
            If Not rsTmp.EOF Then
                .txtInfo(GC_��ԭѧ���).Text = rsTmp!������� & ""
                .cmdInfo(GC_��ԭѧ���).Tag = Val(rsTmp!����id & "")
                Call UpdateCacheRecInfo(0, "��ԭѧ���", rsTmp!������� & rsTmp!����id & "", , , Val(rsTmp!��¼��Դ & ""))
            End If

            Set rsTmp = GetOPSData(lng����ID, lng��ҳID, Not gclsPros.Is��Ŀ, gclsPros.Moved)
            rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
            If gclsPros.FuncType = f������ҳ And rsTmp.EOF Then
                rsTmp.Filter = "��¼��Դ=3"
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1"
                End If
            End If
            '��������
            Call CacheLoadVsOPSData(.vsOPS, rsTmp)
            '��Ϸ���������أ���Ϸ����������������ϣ�ʬ���־�йأ���˷������
            Call CacheLoadDiagMatchData(GetDiagMatchData(lng����ID, lng��ҳID))
            '����ҩʹ���������(�ӱ���Ϣ��Ҳ���ڿ���ҩ��������ݣ������ݣ���������ӱ���Ϣ�ۺϼ���)
            Call CacheLoadVsKSSData(.vsKSS, GetKSSData(lng����ID, lng��ҳID))
            '��֢�໤ʹ���������
            If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                Call CacheLoadVsFlxAddICUData(, GetICUData(lng����ID, lng��ҳID))
            ElseIf gclsPros.MedPageSandard = ST_��������׼ Or gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Call CacheLoadVsFlxAddICUData(.vsFlxAddICU, GetICUData(lng����ID, lng��ҳID))
            End If
            '��֢�໤��еʹ�á�ҽԺ��Ⱦ���걾���
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Call CacheLoadVsICUInstrumentsData(.vsICUInstruments, GetICUInstrumentsData(lng����ID, lng��ҳID))
                Call CacheLoadvsInfectData(.vsInfect, GetInfectData(lng����ID, lng��ҳID))
                Call CacheLoadvsSampleData(.vsSample, GetSampleData(lng����ID, lng��ҳID))
            End If
            '���ơ����ơ�����ҩƷ����(�����������ڲ������׼��ϵͳ����ʱ�ż���)
            If gclsPros.ReadPages Then
                Call CacheLoadVsChemothData(.vsChemoth, GetChemothData(lng����ID, lng��ҳID))
                Call CacheLoadVsRadiothData(.vsRadioth, GetRadiothData(lng����ID, lng��ҳID))
                If gclsPros.MedPageSandard = ST_��������׼ Then
                    Call CacheLoadVsSpiritData(.vsSpirit, GetSpiritData(lng����ID, lng��ҳID))
                End If
            End If
            Call GetDaysFromLast
        End If
        Call SetAllVSF
        
         
    '����Ƿ�����Ҳ���������ҳ
    Call CreatePlugInOK(gclsPros.Module)
    If Not gobjPlugIn Is Nothing Then
        Err.Clear: On Error Resume Next
        If gobjPlugIn.gblnLoadMec = True Then
            '���ò����Զ�����ؽӿ�
            If Err.Number = 0 Then
                Set gColCtl = CtlAdd
                Call gobjPlugIn.LoadMecInfo(gclsPros.SysNo, gclsPros.Module, lng����ID, lng��ҳID, gclsPros.PatiType, gColCtl)
            End If
            Call zlPlugInErrH(Err, "LoadMecInfo")
            Err.Clear: On Error GoTo 0
        End If
    End If
    

    '���ò�����ҳ��Ҳ��������Զ��帽ҳ����
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Err.Clear: On Error Resume Next
            Call gfrmMecCol(i).LoadPlugMec(gclsPros.SysNo, gclsPros.Module, lng����ID, lng��ҳID, gclsPros.PatiType)
            Call zlPlugInErrH(Err, "LoadPlugMec")
            Err.Clear: On Error GoTo 0
        Next
    End If
    
    End With
    Screen.MousePointer = 0
    LoadMedPageData = True
    Exit Function
errH:
    Debug.Print "LoadMedPageData:" & Err.Source & "===" & Err.Description
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'��ʼ������ؼ�����
Public Function InitMedRecEnv(Optional ByVal blnAfterLoadData As Boolean, Optional ByVal blnReLoad As Boolean) As Boolean
'���ܣ���ʼ����ҳ�༭ʱ����Ҫ��һЩ����
'������blnAfterLoadData=�Ƿ������ݼ���֮���ʼ����True-�����ݼ���֮���ʼ����False-�����ݼ���֮ǰ��ʼ��
'      blnReLoad=�Ƿ������³�ʼ��
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strSql As String, i As Long
    Dim vsTmp As VSFlexGrid, LngCol As Long, LngRow As Long
    Dim objTextBox As TextBox, objCmd As CommandButton, objPadr As PatiAddress
    Dim bln�������� As Boolean
    Dim datCur As Date

    On Error GoTo errH
    If Not blnAfterLoadData Then
        With gclsPros.CurrentForm
            Screen.MousePointer = 11
            '���ò���������ĸ߶ȵĿ��
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_ְҵ).hwnd, .cboBaseInfo(BCC_ְҵ).Width + 2800)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_����).hwnd, .cboBaseInfo(BCC_����).Width + 1600)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_����).hwnd, .cboBaseInfo(BCC_����).Width + 800)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_��ϵ).hwnd, .cboBaseInfo(BCC_��ϵ).Width + 1000)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_�ֻ��̶�).hwnd, .cboBaseInfo(BCC_�ֻ��̶�).Width + 500)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_����������).hwnd, .cboBaseInfo(BCC_����������).Width + 1200)
            
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_����), .cboBaseInfo(BCC_����).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_����), .cboBaseInfo(BCC_����).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_����), .cboBaseInfo(BCC_����).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_ְҵ), .cboBaseInfo(BCC_ְҵ).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_���ʽ), .cboBaseInfo(BCC_���ʽ).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_��ϵ), .cboBaseInfo(BCC_��ϵ).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_�ֻ��̶�), .cboBaseInfo(BCC_�ֻ��̶�).Height * 16)
            
            If gclsPros.PatiType <> PF_���� Then
                Call zlControl.CboSetHeight(.cboManInfo(MC_����ҽʦ), .cboManInfo(MC_����ҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_����ҽʦ).hwnd, .cboManInfo(MC_����ҽʦ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_������), .cboManInfo(MC_������).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_������).hwnd, .cboManInfo(MC_������).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_���λ�����), .cboManInfo(MC_���λ�����).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_���λ�����).hwnd, .cboManInfo(MC_���λ�����).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_����ҽʦ), .cboManInfo(MC_����ҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_����ҽʦ).hwnd, .cboManInfo(MC_����ҽʦ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_סԺҽʦ), .cboManInfo(MC_סԺҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_סԺҽʦ).hwnd, .cboManInfo(MC_סԺҽʦ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_����ҽʦ), .cboManInfo(MC_����ҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_����ҽʦ).hwnd, .cboManInfo(MC_����ҽʦ).Width + 1600)
                If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                    Call zlControl.CboSetHeight(.cboManInfo(MC_����ҽʦ), .cboManInfo(MC_����ҽʦ).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_����ҽʦ).hwnd, .cboManInfo(MC_����ҽʦ).Width + 1600)
                Else
                    Call zlControl.CboSetHeight(.cboManInfo(MC_�о���ҽʦ), .cboManInfo(MC_�о���ҽʦ).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_�о���ҽʦ).hwnd, .cboManInfo(MC_�о���ҽʦ).Width + 1600)
                End If
                
                If gclsPros.FuncType = f������ҳ Then
                    Call zlControl.CboSetHeight(.cboManInfo(MC_��ĿԱ), .cboManInfo(MC_��ĿԱ).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_��ĿԱ).hwnd, .cboManInfo(MC_��ĿԱ).Width + 1600)
                End If
                Call zlControl.CboSetHeight(.cboManInfo(MC_ʵϰҽʦ), .cboManInfo(MC_ʵϰҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_ʵϰҽʦ).hwnd, .cboManInfo(MC_ʵϰҽʦ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_�ʿ�ҽʦ), .cboManInfo(MC_�ʿ�ҽʦ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_�ʿ�ҽʦ).hwnd, .cboManInfo(MC_�ʿ�ҽʦ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_�ʿػ�ʿ), .cboManInfo(MC_�ʿػ�ʿ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_�ʿػ�ʿ).hwnd, .cboManInfo(MC_�ʿػ�ʿ).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_���λ�ʿ), .cboManInfo(MC_���λ�ʿ).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_���λ�ʿ).hwnd, .cboManInfo(MC_���λ�ʿ).Width + 1600)
            End If

'            ���̶ֹ����ݵ�������
            Call SetCboFromList(Array("", "0-δ����", "1-����1̥", "2-����2̥������", "4-����"), Array(.cboBaseInfo(BCC_����״��)), 0)
            Call SetCboFromList(Array("-", "0-δ��", "1-��", "2-��", "3-����"), Array(.cboBaseInfo(BCC_RH)))
            Call SetCboFromList(Array("��", "��", "��", "Сʱ", "����"), Array(.cboSpecificInfo(SLC_����)), 0) '�����Ŀʱ��ע��cboInfo(cbo���䵥λ).listIndex<3���ж�
            If gclsPros.PatiType <> PF_���� Then
                Call SetCboFromList(Array("��", "��", "��", "��", "����"), Array(.cboSpecificInfo(SLC_��������)), 0)
                Call SetCboFromList(Array("1.1-��", "1.2-����", "2-����", "3-��"), Array(.cboBaseInfo(BCC_�������), .cboBaseInfo(BCC_���ȷ���)))
                Call SetCboFromList(Array("0-δ֪", "1-��", "2-��"), Array(.cboBaseInfo(BCC_������ҩ�Ƽ�)))
                Call SetCboFromList(Array(" ", "1-��", "2-��"), Array(.cboBaseInfo(BCC_��ҽ�����豸), .cboBaseInfo(BCC_��ҽ���Ƽ���), .cboBaseInfo(BCC_��֤ʩ��)))
                Call SetCboFromList(Array("0-δ��", "1-׼ȷ", "2-����׼ȷ", "3-�ش�ȱ��", "4-����"), Array(.cboBaseInfo(BCC_��֤), .cboBaseInfo(BCC_�η�), .cboBaseInfo(BCC_��ҩ)))
                Call SetCboFromList(Array("1-��", "2-��", "3-δ��", "4-��ȷ��"), Array(.cboBaseInfo(BCC_��Һ��Ӧ)))
                Call SetCboFromList(Array("0-��", "1-��", "2-δ��", "3-��ȷ��"), Array(.cboBaseInfo(BCC_��Ѫ��Ӧ)))
                Call SetCboFromList(Array("-"), Array(.cboBaseInfo(BCC_��������ʬ��)), 0)
                Call SetCboFromList(Array("-"), Array(.cboBaseInfo(BCC_�ٴ���ʬ��)), 0)
                Call SetCboFromList(Array("1-��", "2-��", "3-����"), Array(.cboBaseInfo(BCC_��Ѫǰ9����)))
                Call SetCboFromList(Array("0-δ��", "1-����", "2-������", "3-���϶�"), Array(.cboBaseInfo(BCC_�������ԺXY), .cboBaseInfo(BCC_��������Ժ), .cboBaseInfo(BCC_��Ժ���ԺXY), .cboBaseInfo(BCC_�����벡��), .cboBaseInfo(BCC_�ٴ��벡��), _
                                                                                            .cboBaseInfo(BCC_��ǰ������), .cboBaseInfo(BCC_�������ԺZY), .cboBaseInfo(BCC_��Ժ���ԺZY)))
                Call SetCboFromList(Array(" ", "0-Ժ��", "1-סԺ�ڼ�"), Array(.cboBaseInfo(BCC_ѹ�������ڼ�)))
                Call SetCboFromList(Array(" ", "1��", "2��", "3��", "4��", "5��", "6��"), Array(.cboBaseInfo(BCC_ѹ������)))
                Call SetCboFromList(Array(" ", "һ��", "����", "����", "δ����˺�"), Array(.cboBaseInfo(BCC_������׹���˺�)))
                Call SetCboFromList(Array(" ", "����ԭ��", "���ơ�ҩ�����ԭ��", "��������", "����ԭ��"), Array(.cboBaseInfo(BCC_������׹��ԭ��)))
                Call SetCboFromList(Array("��", "��", "Сʱ", "����"), Array(.cboSpecificInfo(SLC_Ӥ�׶�����)), 0) '�����Ŀʱ��ע��cboInfo(cbo���䵥λ).listIndex<3���ж�
                Call SetCboFromList(Array("31������סԺ�ƻ�", "7������סԺ�ƻ�"), Array(.cboBaseInfo(BCC_����Ժ�ƻ�����)), 0)
                Call SetCboFromList(Array("", "1-��", "2-��", "3-��"), Array(.cboBaseInfo(BCC_��������)))
                Call SetCboFromList(Array("", "0-ֱ��", "1-���", "2-��"), Array(.cboBaseInfo(BCC_��Ⱦ��������ϵ)), 0)

                If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                    Call SetCboFromList(Array("0-δ��", "1-����", "2-����", "3-������"), Array(.cboBaseInfo(BCC_HBsAg)))
                    Call SetCboFromList(Array("0-δ��", "1-����", "2-����", "3-��ȷ��"), Array(.cboBaseInfo(BCC_HCVAb), .cboBaseInfo(BCC_HIVAb)))
                End If

                If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    If gclsPros.FuncType = f������ҳ Then
                        Call SetCboFromList(Array("��һ��ס��Ժ", "����", "2-15��", "16-31��", "��31��"), Array(.cboBaseInfo(BCC_���ϴ�סԺʱ��)))
                    End If
                    Call SetCboFromList(Array("���ط�", "24h��", "24-48h", "��48h"), Array(.cboBaseInfo(BCC_�ط����ʱ��)))
                    Call SetCboFromList(Array("", "һ��", "����", "����", "����"), Array(.cboBaseInfo(BCC_Լ����ʽ)))
                    Call SetCboFromList(Array("", "��ʽ��", "Ӳʽ��", "����", "������", "Լ����", "����"), Array(.cboBaseInfo(BCC_Լ������)))
                    Call SetCboFromList(Array("", "��֪�ϰ�", "���ܵ���", "��Ϊ����", "������Ҫ", "�궯", "ҽ������", "����"), Array(.cboBaseInfo(BCC_Լ��ԭ��)))
                    Call SetCboFromList(Array("", "ҽ����Ժ", "ת����", "תԺ", "��ҽ����Ժ", "����"), Array(.cboBaseInfo(BCC_��������Ժ��ʽ)))
                ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    Call SetCboFromList(Array("", "1-δ����", "2-�����˳�", "3-���"), Array(.cboBaseInfo(BCC_�ٴ�·������)), 0)
                    Call SetCboFromList(Array("", "1-��", "2-������", "3-������", "4-���߶���"), Array(.cboBaseInfo(BCC_ʵʩDGRS����)), 0)
                    Call SetCboFromList(Array("", "1-����", "2-����", "3-����"), Array(.cboBaseInfo(BCC_������Ⱦ��)), 0)
                    Call SetCboFromList(Array("", "1-0��", "2-I��", "3-����", "4-����", "5-����", "6-����"), Array(.cboBaseInfo(BCC_��������)), 0)
                ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                    Call SetCboFromList(Array("", "1-0��", "2-I��", "3-����", "4-����", "5-����", "6-����"), Array(.cboBaseInfo(BCC_��������)), 0)
                End If
            End If
            '����һЩ�ֵ���������������
            Call SetCboFromRec(Array(BCC_���ʽ, BCC_�Ա�, BCC_����, BCC_ְҵ, BCC_����, BCC_����, BCC_Ѫ��, BCC_���֤), 0)
            If gclsPros.PatiType <> PF_���� Then
                Call SetCboFromRec(Array(BCC_��������), 0, "")
                Call SetCboFromRec(Array(BCC_��ϵ, BCC_��Ժ���, BCC_��Ժ;��, BCC_�ֻ��̶�, BCC_����������, BCC_��Ժ��ʽ), 0)
                '��Ⱦ��λ����Ⱦ���أ������¼���ListBOX�ļ���
                Call SetLstBoxFromRec(IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, "��Ⱦ��λ,�����¼�", "��Ⱦ��λ,��Ⱦ����,�����¼�"))
            Else
                Call SetCboFromRec(Array(BCC_ȥ��), 0, " ")
                Call SetCboFromRec(Array(BCC_�Ļ��̶�), 0)
            End If
            Call SetCboFromRec(Array(BCC_�����ڼ�), 0)
            If gclsPros.FuncType = f������ҳ Then
                '�õ�Ĭ�ϳ�����
                strSql = "select A.����,A.���� from ���� a where a.ȱʡ��־=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_�����ص�, "�����ص�", rsTmp!����, True)
                    If gclsPros.DefautADD Then
                        Call SetPatiAddress(ADRC_��ϵ�˵�ַ, "��ϵ�˵�ַ", rsTmp!����, True)
                        Call SetPatiAddress(ADRC_��סַ, "��ͥ��ַ", rsTmp!����, True)
                        .txtSpecificInfo(SLC_��ͥ�ʱ�).Text = rsTmp!���� & ""
                    End If
                End If
                '����:13557
                strSql = "select A.����,A.���� from ���� a where a.ȱʡ��־=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_��������, "����", rsTmp!����, True)
                End If
                datCur = zlDatabase.Currentdate
                '����Ĭ��ֵ
                .mskDateInfo(DC_��Ŀ����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ŀ����).Tag))
                .mskDateInfo(DC_�ջ�����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_�ջ�����).Tag))
                .mskDateInfo(DC_��������).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��������).Tag))
                .mskDateInfo(DC_��Ժʱ��).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ժʱ��).Tag))
                .mskDateInfo(DC_��Ժʱ��).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ժʱ��).Tag))
                .mskDateInfo(DC_�ʿ�����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_�ʿ�����).Tag))

                .txtDateInfo(DC_��Ŀ����).Text = .mskDateInfo(DC_��Ŀ����).Text
                .txtDateInfo(DC_�ջ�����).Text = .mskDateInfo(DC_�ջ�����).Text
                .txtDateInfo(DC_��������).Text = .mskDateInfo(DC_��������).Text
                .txtDateInfo(DC_��Ժʱ��).Text = .mskDateInfo(DC_��Ժʱ��).Text
                .txtDateInfo(DC_��Ժʱ��).Text = .mskDateInfo(DC_��Ժʱ��).Text
                .txtDateInfo(DC_�ʿ�����).Text = .mskDateInfo(DC_�ʿ�����).Text

                gclsPros.InTime = .mskDateInfo(DC_��Ժʱ��).Text
                gclsPros.OutTime = .mskDateInfo(DC_��Ժʱ��).Text
                '������������
                .cmdFeeEdit.Visible = Not gclsPros.OnLine And (gclsPros.OpenMode = EM_�������� Or gclsPros.OpenMode = EM_������ҳ) And gclsPros.OutFile <> ""
            End If
            If gclsPros.PatiType <> PF_���� Then
                '������ز����Լ���������ʼ������
                '�ṹ����ַ
                On Error Resume Next
                For Each objTextBox In .txtAdressInfo
                    Set objPadr = .padrInfo(objTextBox.Index)
                    strTmp = objPadr.Name
                    If Err.Number = 0 Then  '���ڵ�ַ�ؼ�����TextBox��CommandButton�����ܲ����ڣ���Ҫ����
                        objTextBox.Visible = Not gclsPros.IsStructAdress
                        Set objCmd = .cmdAdressInfo(objTextBox.Index) '���ܲ����ڣ���Ϊ������ԣ�����ֱ����������
                        objCmd.Visible = Not gclsPros.IsStructAdress
                        objPadr.Visible = gclsPros.IsStructAdress
                        objPadr.ShowTown = gclsPros.IsShowTown
                    Else
                        Err.Clear
                    End If
                Next
                If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    Call SetCboFromRec(Array(BCC_����ԭ��), 0)
                    .cboBaseInfo(BCC_����ԭ��).Visible = gclsPros.PathVCauses
                    .fraCbo(0).Visible = gclsPros.PathVCauses
                    .txtInfo(GC_����ԭ��).Visible = Not gclsPros.PathVCauses
                End If

                On Error GoTo errH
                '�Ĵ����ȡ�ϴ���ϡ��������ϻ�ȡ�ϴ����
                If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ And gclsPros.FuncType = f������ҳ Then
                    .cmdLastDiag.Visible = gclsPros.��ҳID > 1
                End If
            End If
        End With
        If Not InitTableDiag Then Exit Function
        If Not InitTableAller Then Exit Function
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
            If Not InitTableFees Then Exit Function
        End If
    Else
        '���ݼ��غ���������
        '����ҽ��ϼ��غ��������
        Set vsTmp = gclsPros.CurrentForm.vsDiagXY
        With vsTmp
            .Cell(flexcpForeColor, 1, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            If gclsPros.PatiType <> PF_���� Then
                LngRow = FindDiagRow(DT_��Ժ���XY)
                .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                .Row = .FixedRows: .Col = DI_�������
                Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
            Else
                .Cell(flexcpText, .FixedRows, DI_�������, .Rows - 1, DI_�������) = "��ҽ"
            End If
        End With

        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            If gclsPros.PatiType <> PF_���� Then
                LngRow = FindDiagRow(DT_��Ժ���ZY)
                .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
            Else
                .Cell(flexcpText, .FixedRows, DI_�������, .Rows - 1, DI_�������) = "��ҽ"
            End If
        End With
        '���ݼ��غ����ÿ���ҩ�������
        If gclsPros.PatiType <> PF_���� Then Call SetKSSSerial
        With gclsPros.CurrentForm
            If gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiType <> PF_���� Then
                '���۲�����סԺ��
                If Val(gclsPros.PatiInfo!�������� & "") <> 0 Then
                    .lblSpecificInfo(SLC_סԺ��).Visible = False
                    .txtSpecificInfo(SLC_סԺ��).Visible = False
                    .txtSpecificInfo(SLC_סԺ��).Enabled = False '��־Ϊ�����
                    .PicInNum.Visible = False
                End If
            End If
            '������ҳ������ҽ���ʱ������ҽ��ϱ��
            If Not gclsPros.Have��ҽ And gclsPros.PatiType = PF_���� Then
                .vsDiagZY.Visible = False
                .vsDiagXY.Height = .vsDiagZY.Top + .vsDiagZY.Height - .vsDiagXY.Top
               .vsDiagXY.ColHidden(DI_��Ϸ���) = True
                .vsDiagXY.ColWidth(DI_��ϱ���) = .vsDiagXY.ColWidth(DI_��ϱ���) + .vsDiagXY.ColWidth(DI_��Ϸ���)
            End If
            '������ҳ(��Ժ���һ���Ժ�����в�������)���Ʋ���������ʿ
            If gclsPros.PatiType <> PF_���� Then .vsOPS.ColHidden(PI_������ʿ) = Not gclsPros.Is����
        End With
    End If
    Screen.MousePointer = 0
    InitMedRecEnv = True
    Exit Function
errH:
    Debug.Print "InitMedRecEnv:" & Err.Source & "===" & Err.Description
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDiagData(ByVal curDate As Date, Optional ByRef blnHaveSel As Boolean) As Boolean
    Dim vsTmp As VSFlexGrid
    Dim lngSize As Long
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim blnHaveDaig As Boolean
    Dim lng��ҽNotHave As Long '��ҽȱʧ���
    Dim lng��ҽNotHave As Long '��ҽȱʧ���
    Dim lngSame As Long '��ҽ��ͬ��ϲ�֤ͬ�������
    Dim lngSameType As Long
    
    On Error GoTo errH
    
    gclsPros.DiagNames = "": gclsPros.DiagRowIDs = ""
    gclsPros.DiseaseIDs = "": gclsPros.DiagIDs = ""
    '������ҳ����ҽ����ҽ����һ����ϵģ������Ժ����Ժ�����ȫ
    If gclsPros.FuncType = f������ҳ Then
        If gclsPros.Have��ҽ Then
            Set vsTmp = gclsPros.CurrentForm.vsDiagZY
            If vsTmp.TextMatrix(FindDiagRow(DT_�������ZY), DI_�������) = "" Then
                lng��ҽNotHave = DT_�������ZY
            End If
            If lng��ҽNotHave = 0 Then
                If vsTmp.TextMatrix(FindDiagRow(DT_��Ժ���ZY), DI_�������) = "" Then
                    lng��ҽNotHave = DT_��Ժ���ZY
                End If
            End If
            If lng��ҽNotHave = 0 Then
                If vsTmp.TextMatrix(FindDiagRow(DT_��Ժ���ZY), DI_�������) = "" Then
                    lng��ҽNotHave = DT_��Ժ���ZY
                End If
            End If
        End If
        Set vsTmp = gclsPros.CurrentForm.vsDiagXY
         If vsTmp.TextMatrix(FindDiagRow(DT_�������XY), DI_�������) = "" Then
             lng��ҽNotHave = DT_�������XY
         End If
         If lng��ҽNotHave = 0 Then
             If vsTmp.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_�������) = "" Then
                 lng��ҽNotHave = DT_��Ժ���XY
             End If
         End If
         If lng��ҽNotHave = 0 Then
             If vsTmp.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_�������) = "" Then
                 lng��ҽNotHave = DT_��Ժ���XY
             End If
         End If
         If lng��ҽNotHave <> 0 And (lng��ҽNotHave <> 0 And gclsPros.Have��ҽ Or Not gclsPros.Have��ҽ) Then
             If gclsPros.Have��ҽ Then
                Set vsTmp = gclsPros.CurrentForm.vsDiagZY
                vsTmp.Row = FindDiagRow(lng��ҽNotHave): vsTmp.Col = DI_�������
                If gclsPros.FuncType = f���ѡ�� Then
                    Call ShowMessage(vsTmp, "��ҽ��ϵ�" & decode(lng��ҽNotHave, DT_�������ZY, "�ţ����������", DT_��Ժ���ZY, "��Ժ���", "��Ժ���") & "����Ҫ��ϲ���Ϊ�ա�")
                    Exit Function
                Else
                    Call AddErrInfo("��ҽ��ϵ�" & decode(lng��ҽNotHave, DT_�������ZY, "�ţ����������", DT_��Ժ���ZY, "��Ժ���", "��Ժ���") & "����Ҫ��ϲ���Ϊ�ա�", 0, vsTmp)
                End If
             Else
                Set vsTmp = gclsPros.CurrentForm.vsDiagXY
                vsTmp.Row = FindDiagRow(lng��ҽNotHave): vsTmp.Col = DI_�������
                If gclsPros.FuncType = f���ѡ�� Then
                    Call ShowMessage(vsTmp, "��ҽ��ϵ�" & decode(lng��ҽNotHave, DT_�������XY, "�ţ����������", DT_��Ժ���XY, "��Ժ���", "��Ժ���") & "����Ҫ��ϲ���Ϊ�ա�")
                    Exit Function
                Else
                    Call AddErrInfo("��ҽ��ϵ�" & decode(lng��ҽNotHave, DT_�������XY, "�ţ����������", DT_��Ժ���XY, "��Ժ���", "��Ժ���") & "����Ҫ��ϲ���Ϊ�ա�", 0, vsTmp)
                End If
             End If
         End If
    End If
    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    'gclsPros.InsureType = 920 And gclsPros.Module = p����ҽ��վ,ԭ����ע���Ǳ���ҽ��������Ҫ��(�����ҵ���)
    lngSize = IIf(gclsPros.InsureType = 920 And gclsPros.PatiType = PF_����, 82, 200)
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, DI_�������)) <> "" Then
                blnHaveDaig = True
                If i <> .Rows - 1 Then '����Ƿ����������ͬ�������ͬ���������
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(.TextMatrix(i, DI_��Ϸ���)) And Val(.TextMatrix(i, DI_��Ϸ���)) <> DT_������� Then
                            If .TextMatrix(j, DI_�������) <> "" Then
                                If .TextMatrix(j, DI_�������) = .TextMatrix(i, DI_�������) Then
                                    .Row = i: .Col = DI_�������
                                    If gclsPros.FuncType = f���ѡ�� Then
                                        Call ShowMessage(vsTmp, "���ִ���������ͬ�������Ϣ��")
                                        Exit Function
                                    Else
                                        If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                            Exit For
                                        Else
                                            Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�з��ִ�����ͬ�������Ϣ��", 0, vsTmp)
                                            lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                            Exit For
                                        End If
                                    End If
                                ElseIf Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                                    If Val(.TextMatrix(j, DI_����ID)) = Val(.TextMatrix(i, DI_����ID)) Then
                                        .Row = i: .Col = DI_�������
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "���ִ���������ͬ�������Ϣ��")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�з��ִ�����ͬ�������Ϣ��", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                ElseIf Val(.TextMatrix(i, DI_���ID)) <> 0 Then
                                    If Val(.TextMatrix(j, DI_���ID)) = Val(.TextMatrix(i, DI_���ID)) Then
                                        .Row = i: .Col = DI_�������
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "���ִ���������ͬ�������Ϣ��")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�з��ִ�����ͬ�������Ϣ��", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                If Val(.TextMatrix(i, DI_��Ϸ���)) = DT_������� Then
                    If gclsPros.FuncType = f������ҳ Or gclsPros.FuncType = fҽ����ҳ Then
                        If Not gclsPros.AddPathologic Then
                            If gclsPros.CurrentForm.txtInfo(GC_�����).Text = "" Then
                                Call AddErrInfo("����ϱ�����д����ţ�����д����š�", 0, gclsPros.CurrentForm.txtInfo(GC_�����))
                            End If
                        End If
                        If gclsPros.CurrentForm.cboBaseInfo(BCC_�ٴ��벡��).Text = "" Then
                            Call AddErrInfo("����ϱ�����д�ٴ��벡������д�ٴ��벡��", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_�ٴ��벡��))
                        End If
                        If gclsPros.CurrentForm.cboBaseInfo(BCC_�����벡��).Text = "" Then
                            Call AddErrInfo("����ϱ�����д�����벡������д�����벡��", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_�����벡��))
                        End If
                    End If
                End If
                If .TextMatrix(i - 1, DI_�������) = "" And Val(.TextMatrix(i, DI_��Ϸ���)) = Val(.TextMatrix(i - 1, DI_��Ϸ���)) Then
                    .Row = i - 1: .Col = DI_�������
                    If gclsPros.FuncType = f���ѡ�� Then
                        Call ShowMessage(vsTmp, "���������������Ϣ��")
                        Exit Function
                    Else
                        Call AddErrInfo("���������������Ϣ��", 0, vsTmp)
                    End If
                End If
                
                If zlCommFun.ActualLen(.TextMatrix(i, DI_�������)) > lngSize Then
                    .Row = i: .Col = DI_�������
                    If gclsPros.FuncType = f���ѡ�� Then
                        Call ShowMessage(vsTmp, .TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "����̫����ֻ����" & lngSize & "���ַ���" & lngSize / 2 & "�����֡�")
                        Exit Function
                    Else
                        Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "����̫����ֻ����" & lngSize & "���ַ���" & lngSize / 2 & "�����֡�", 0, vsTmp)
                    End If
                End If
                If gclsPros.PatiType = PF_���� Then
                    If .TextMatrix(i, DI_����ʱ��) <> "" Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, DI_����ʱ��), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = DI_����ʱ��
                            If gclsPros.FuncType = f���ѡ�� Then
                                Call ShowMessage(vsTmp, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                                Exit Function
                            Else
                                Call AddErrInfo("����ʱ��Ӧ�����ڵ�ǰʱ�䡣", 0, vsTmp)
                            End If
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.TextMatrix(i, DI_��ע)) > 200 Then
                        .Row = i: .Col = DI_��ע
                        If gclsPros.FuncType = f���ѡ�� Then
                            Call ShowMessage(vsTmp, """" & .TextMatrix(i, DI_�������) & """�ı�ע����̫����ֻ����200���ַ���100�����֡�")
                            Exit Function
                        Else
                            Call AddErrInfo("""" & .TextMatrix(i, DI_�������) & """�ı�ע����̫����ֻ����200���ַ���100�����֡�", 0, vsTmp)
                        End If
                    End If
                    If gclsPros.FuncType = f������ҳ Then
                        If .TextMatrix(i, DI_����ID) = .TextMatrix(i, DI_����ID) And Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                            .Row = i: .Col = DI_��ϱ���
                            If gclsPros.FuncType = f���ѡ�� Then
                                Call ShowMessage(vsTmp, "ͬһ����ϵ������븽�벻����ͬ��")
                                Exit Function
                            Else
                                Call AddErrInfo("ͬһ����ϵ������븽�벻����ͬ��", 0, vsTmp)
                            End If
                        End If
                        If .TextMatrix(i, DI_��ϱ���) = "" And (Not gclsPros.CNIndent Or Val(.TextMatrix(i, DI_��Ϸ���)) <> DT_��Ժ���XY) Then
                            If Val(.TextMatrix(i, DI_��Ϸ���)) <> DT_������� Then
                                .Row = i: .Col = DI_��ϱ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "����ϱ�����д��ϱ��룬�������б������ϻ���ϱ��롣")
                                    Exit Function
                                Else
                                    Call AddErrInfo("����ϱ�����д��ϱ��룬�������б������ϻ���ϱ��롣", 0, vsTmp)
                                End If
                            End If
                        End If
                    End If
                    
                    If .TextMatrix(i, DI_��Ч����) <> "" And InStr(.TextMatrix(i, DI_��Ժ���), .TextMatrix(i, DI_��Ч����)) > 0 Then
                        If gclsPros.FuncType = f���ѡ�� Then
                            If ShowMessage(vsTmp, "��" & .TextMatrix(i, DI_�������) & "�������ĳ�Ժ���Ϊ��" & .TextMatrix(i, DI_��Ժ���) & "��" & _
                                vbCrLf & "�Ƿ�ȷ�ϣ�", True) = vbNo Then
                                .Row = i: .Col = DI_��Ժ���
                                Exit Function
                            End If
                        Else
                            .Row = i: .Col = DI_��Ժ���
                            Call AddErrInfo("��" & .TextMatrix(i, DI_�������) & "�������ĳ�Ժ���Ϊ��" & .TextMatrix(i, DI_��Ժ���) & "���Ƿ�ȷ�ϣ�", 1, vsTmp)
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_��Ϸ���)) = DT_Ժ�ڸ�Ⱦ Then
                        If .TextMatrix(i, DI_��Ժ���) = "" Then
                            If gclsPros.FuncType = f������ҳ Then
                                If Not gclsPros.Null��Ժ��� Then
                                    .Row = i: .Col = DI_��Ժ���
                                    If gclsPros.FuncType = f���ѡ�� Then
                                        Call ShowMessage(vsTmp, "����дԺ�ڸ�Ⱦ�ĳ�Ժ�����")
                                        Exit Function
                                    Else
                                        Call AddErrInfo("����дԺ�ڸ�Ⱦ�ĳ�Ժ�����", 0, vsTmp)
                                    End If
                                End If
                            Else
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    If ShowMessage(vsTmp, "Ժ�ڸ�Ⱦ�ĳ�Ժ���û����д���Ƿ������", True) = vbNo Then Exit Function
                                Else
                                    Call AddErrInfo("Ժ�ڸ�Ⱦ�ĳ�Ժ���û����д���Ƿ������", 1, vsTmp)
                                End If
                            End If
                        End If
                    ElseIf Val(.TextMatrix(i, DI_��Ϸ���)) = DT_��Ժ���XY Then
                        If .TextMatrix(i, DI_��Ժ����) = "" And DiagCellEditable(vsTmp, i, DI_��Ժ����) Then
                            .Row = i: .Col = DI_��Ժ����
                            If gclsPros.FuncType = f���ѡ�� Then
                                Call ShowMessage(vsTmp, "����д��Ժ���顣")
                                Exit Function
                            Else
                                Call AddErrInfo("����д��Ժ���顣", 0, vsTmp)
                            End If
                        End If
                        
                        If .TextMatrix(i, DI_��Ժ���) = "" Then
                            If gclsPros.FuncType = f������ҳ Then
                                If Not gclsPros.Null��Ժ��� Then
                                    .Row = i: .Col = DI_��Ժ���
                                    If gclsPros.FuncType = f���ѡ�� Then
                                        Call ShowMessage(vsTmp, "�����Ժ�����")
                                        Exit Function
                                    Else
                                        Call AddErrInfo("�����Ժ�����", 0, vsTmp)
                                    End If
                                End If
                            Else
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "����д��Ժ�����")
                                    Exit Function
                                Else
                                    Call AddErrInfo("�����Ժ�����", 0, vsTmp)
                                End If
                            End If
                        End If
                        If .TextMatrix(i, DI_�������) <> "��Ժ���" Then
                            If InStr(.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_��Ժ���), "����") = 0 And InStr(.TextMatrix(i, DI_��Ժ���), "����") > 0 Then
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    If InStr(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), "����") = 0 Then
                                        Call ShowMessage(vsTmp, "��Ҫ��ϵĳ�Ժ�����Ϊ����������������ϵĳ�Ժ���ȴΪ������")
                                        Exit Function
                                    End If
                                Else
                                    If InStr(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), "����") = 0 Then
                                        Call AddErrInfo("��Ҫ��ϵĳ�Ժ�����Ϊ����������������ϵĳ�Ժ���ȴΪ������", 0, vsTmp)
                                    End If
                                End If
                            End If
                        Else '��Ҫ��Ժ���
                            If InStr(.TextMatrix(i, DI_��Ժ���), "����") > 0 And gclsPros.Have���� Then
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    If ShowMessage(vsTmp, "�ò��˽���������������Ժ���ѡ��Ϊ�������Ƿ������", True) = vbNo Then Exit Function
                                Else
                                    Call AddErrInfo("�ò��˽���������������Ժ���ѡ��Ϊ�������Ƿ������", 1, vsTmp)
                                End If
                            End If
                            If gclsPros.FuncType <> f���ѡ�� Then
                                If InStr(.TextMatrix(i, DI_��Ժ���), "����") > 0 And Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ����).Text) < 3 Then
                                    .Row = i: .Col = DI_��Ժ���
                                    If gclsPros.FuncType = f���ѡ�� Then
                                        If ShowMessage(vsTmp, "�ò���סԺ��ԺΪ " & Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ����).Text) & " �죬��Ժ���ȴΪ�������Ƿ������", True) = vbNo Then Exit Function
                                    Else
                                        Call AddErrInfo("�ò���סԺ��ԺΪ " & Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ����).Text) & " �죬��Ժ���ȴΪ�������Ƿ������", 1, vsTmp)
                                    End If
                                End If
                            End If
                            If gclsPros.Check�����ж� <> 0 Then
                                '��Ҫ�����Ҫ�����˵��ⲿԭ��
                                If InStr("ST", Left(.TextMatrix(i, DI_��ϱ���), 1)) > 0 And Left(.TextMatrix(i, DI_��ϱ���), 1) <> "" Then
                                    '��Ҫ�����ж��ⲿԭ��
                                    If .TextMatrix(FindDiagRow(DT_�����ж���), DI_�������) = "" Then
                                        .Row = FindDiagRow(DT_�����ж���): .Col = DI_�������
                                        If gclsPros.Check�����ж� = 1 Then
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "����д�����ж���ԭ��")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("����д�����ж���ԭ��", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                If ShowMessage(vsTmp, "û����д�����ж���ԭ��,�Ƿ������", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("û����д�����ж���ԭ��,�Ƿ������", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(FindDiagRow(DT_�����ж���), DI_�������) <> "" Then
                                        .Row = FindDiagRow(DT_�����ж���): .Col = DI_�������
                                        If gclsPros.Check�����ж� = 1 Then
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "������д�����ж���ԭ��")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("������д�����ж���ԭ��", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                If ShowMessage(vsTmp, "��Ժ����������ж���ԭ�򲻷�,�Ƿ������", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("��Ժ����������ж���ԭ�򲻷�,�Ƿ������", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            If gclsPros.Check������� <> 0 Then
                                '��Ҫ�����Ҫ��д������ϵ��ⲿԭ��
                                If (InStr("C", Left(.TextMatrix(i, DI_��ϱ���), 1)) > 0 Or (InStr("D", Left(.TextMatrix(i, DI_��ϱ���), 1)) > 0 And Val(Mid(.TextMatrix(i, DI_��ϱ���), 2, 2)) <= 48)) And Left(.TextMatrix(i, DI_��ϱ���), 1) <> "" Then
                                    '��Ҫ������ϵ��ⲿԭ��
                                    If .TextMatrix(FindDiagRow(DT_�������), DI_�������) = "" Then
                                        .Row = FindDiagRow(DT_�������): .Col = DI_�������
                                        If gclsPros.Check������� = 1 Then
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "��Ժ���Ϊ�������,����д������ϡ�")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("��Ժ���Ϊ�������,����д������ϡ�", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                If ShowMessage(vsTmp, "��Ժ���Ϊ�������,û����д�������,�Ƿ������", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("��Ժ���Ϊ�������,û����д�������,�Ƿ������", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(FindDiagRow(DT_�������), DI_�������) <> "" Then
                                        .Row = FindDiagRow(DT_�������): .Col = DI_�������
                                        If gclsPros.Check������� = 1 Then
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "������д������ϡ�")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("������д������ϡ�", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                If ShowMessage(vsTmp, "��Ժ����벡����ϲ���,�Ƿ������", True) = vbNo Then Exit Function
                                            Else
                                                 Call AddErrInfo("��Ժ����벡����ϲ���,�Ƿ������", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        If (InStr("C", Left(.TextMatrix(i, DI_��ϱ���), 1)) > 0 Or (InStr("D", Left(.TextMatrix(i, DI_��ϱ���), 1)) > 0 And Val(Mid(.TextMatrix(i, DI_��ϱ���), 2, 2)) <= 48)) And Left(.TextMatrix(i, DI_��ϱ���), 1) <> "" Then
                            'ICD�����Ƿ������д
                            If gclsPros.CheckICD���� <> 0 Then
                                .Row = i: .Col = DI_ICD����
                                If .TextMatrix(i, DI_ICD����) = "" Then
                                    If gclsPros.CheckICD���� = 1 Then
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "��ǰ���Ϊ�������,����д������̬ѧ���롣")
                                            Exit Function
                                        Else
                                            Call AddErrInfo("��ǰ���Ϊ�������,����д������̬ѧ���롣", 0, vsTmp)
                                        End If
                                    Else
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            If ShowMessage(vsTmp, "��ǰ���Ϊ�������,û����д������̬ѧ����,�Ƿ������", True) = vbNo Then Exit Function
                                        Else
                                            Call AddErrInfo("��ǰ���Ϊ�������,û����д������̬ѧ����,�Ƿ������", 1, vsTmp)
                                        End If
                                    End If
                                Else
                                    If Left(.TextMatrix(i, DI_ICD����), 1) <> "M" Then

                                        If gclsPros.CheckICD���� = 1 Then
                                            'ICD���������M��ͷ��
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "��ǰ���Ϊ�������,ֻ������д������̬ѧ����(M)��")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("��ǰ���Ϊ�������,ֻ������д������̬ѧ����(M)��", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                If ShowMessage(vsTmp, "��ǰ���Ϊ�������,ֻ������д������̬ѧ����(M),�Ƿ������", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("��ǰ���Ϊ�������,ֻ������д������̬ѧ����(M),�Ƿ������", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                            
                            
  
                If Val(.TextMatrix(i, DI_����ID)) <> 0 Then gclsPros.DiseaseIDs = gclsPros.DiseaseIDs & "," & Val(.TextMatrix(i, DI_����ID))
                If Val(.TextMatrix(i, DI_���ID)) <> 0 Then gclsPros.DiagIDs = gclsPros.DiagIDs & "," & Val(.TextMatrix(i, DI_���ID))
                '�Ƿ�������Ҫ����������
                If gclsPros.PatiType = PF_���� Then
                    gclsPros.IsDiagInput = True
                Else
                    If InStr("," & gclsPros.MustDiagType & ",", "," & Val(.TextMatrix(i, DI_��Ϸ���)) & ",") > 0 Then
                        gclsPros.IsDiagInput = True
                    End If
                End If
                If Val(.TextMatrix(i, DI_����)) <> 0 Then
                    gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs <> "", ",", "") & .RowData(i)
                    strTmp = IIf(Trim(.TextMatrix(i, DI_��ϱ���)) = "", "", "(" & .TextMatrix(i, DI_��ϱ���) & ")") & .TextMatrix(i, DI_�������)
                    gclsPros.DiagNames = gclsPros.DiagNames & IIf(gclsPros.DiagNames <> "", ",", "") & strTmp
                    .Cell(flexcpData, i, DI_����) = ""
                    blnHaveSel = True
                End If
            End If
        Next
    End With

    If gclsPros.Have��ҽ Then
        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, DI_�������) <> "" Then
                    blnHaveDaig = True
                    lngSame = 0
                    lngSameType = 0
                    If i <> .Rows - 1 Then '����Ƿ����������ͬ�������ͬ���������
                        For j = i + 1 To .Rows - 1
                            If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                If Trim(.TextMatrix(j, DI_�������)) <> "" Then
                                    If .TextMatrix(j, DI_�������) & "|" & .TextMatrix(j, DI_��ҽ֤��) = .TextMatrix(i, DI_�������) & "|" & .TextMatrix(i, DI_��ҽ֤��) Then
                                        .Row = i: .Col = DI_�������
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "���ִ���������ͬ�������Ϣ��")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�з��ִ�����ͬ�������Ϣ��", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                                Exit For
                                            End If
                                        End If
                                    ElseIf Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                                        If Val(.TextMatrix(j, DI_����ID)) & "|" & .TextMatrix(j, DI_��ҽ֤��) = Val(.TextMatrix(i, DI_����ID)) & "|" & .TextMatrix(i, DI_��ҽ֤��) Then
                                            .Row = i: .Col = DI_�������
                                            If gclsPros.FuncType = f���ѡ�� Then
                                                Call ShowMessage(vsTmp, "���ִ���������ͬ�������Ϣ��")
                                                Exit Function
                                            Else
                                                If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                                    Exit For
                                                Else
                                                    Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�з��ִ�����ͬ�������Ϣ��", 0, vsTmp)
                                                    lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                    If .TextMatrix(j, DI_�������) = .TextMatrix(i, DI_�������) Then
                                        lngSame = lngSame + 1
                                    ElseIf Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                                        If Val(.TextMatrix(j, DI_����ID)) = Val(.TextMatrix(i, DI_����ID)) Then
                                            lngSame = lngSame + 1
                                        End If
                                    End If
                                    If lngSame >= 2 Then
                                        .Row = i: .Col = DI_�������
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "�����������ϵ������ͬ��֤��ͬ����ϣ���ϲ���ȷ��")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_��Ϸ���)) Then
                                                Exit For
                                            Else
'                                                Call AddErrInfo("�����������ϵ������ͬ��֤��ͬ����ϣ���ϲ���ȷ��", 0, vsTmp)
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�д����������ϵ������ͬ��֤��ͬ����ϣ���ϲ���ȷ��", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_��Ϸ���))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                    If i <> 0 Then
                        If .TextMatrix(i - 1, DI_�������) = "" And Val(.TextMatrix(i, DI_��Ϸ���)) = Val(.TextMatrix(i - 1, DI_��Ϸ���)) Then
                            .Row = i - 1: .Col = DI_�������
                            If gclsPros.FuncType = f���ѡ�� Then
                                Call ShowMessage(vsTmp, "���������������Ϣ��")
                                Exit Function
                            Else
                                Call AddErrInfo("���������������Ϣ��", 0, vsTmp)
                            End If
                        End If
                    End If
                    
                    If zlCommFun.ActualLen(.TextMatrix(i, DI_�������)) + zlCommFun.ActualLen(.TextMatrix(i, DI_��ҽ֤��)) > lngSize Then
                        .Row = i: .Col = DI_�������
                        If gclsPros.FuncType = f���ѡ�� Then
                            Call ShowMessage(vsTmp, .TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�������������ҽ֤������̫���������������ҽ֤�������ֻ����" & lngSize & "���ַ���" & lngSize / 2 & "�����֡�")
                            Exit Function
                        Else
                            Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_�������) = "", FindDiagRow(Val(.TextMatrix(i, DI_��Ϸ���))), i), DI_�������) & "�������������ҽ֤������̫���������������ҽ֤�������ֻ����" & lngSize & "���ַ���" & lngSize / 2 & "�����֡�", 0, vsTmp)
                        End If
                    End If
                    
                    If gclsPros.PatiType = PF_���� Then
                        If .TextMatrix(i, DI_����ʱ��) <> "" Then
                            If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, DI_����ʱ��), "YYYY-MM-DD HH:mm") Then
                                 .Row = i: .Col = DI_����ʱ��
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                                    Exit Function
                                Else
                                    Call AddErrInfo("����ʱ��Ӧ�����ڵ�ǰʱ�䡣", 0, vsTmp)
                                End If
                            End If
                        End If
                        
                         '��ҽ��Ϻ���ҽ��ϵ�����¼��ҽ�����ܴ�����ͬ��
                        If .TextMatrix(i, DI_��ϱ���) = "" Then
                            For j = gclsPros.CurrentForm.vsDiagXY.FixedRows To gclsPros.CurrentForm.vsDiagXY.Rows - 1
                                If gclsPros.CurrentForm.vsDiagXY.TextMatrix(j, DI_��ϱ���) = "" Then
                                    If gclsPros.CurrentForm.vsDiagXY.TextMatrix(j, DI_�������) = .TextMatrix(i, DI_�������) Then
                                        .Row = i: .Col = DI_�������
                                        If gclsPros.FuncType = f���ѡ�� Then
                                            Call ShowMessage(vsTmp, "���ִ���������ͬ������¼�������Ϣ(��ҽ�������ҽ���)��")
                                            Exit Function
                                        Else
                                            Call AddErrInfo("���ִ���������ͬ������¼�������Ϣ(��ҽ�������ҽ���)��", 0, vsTmp)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Else
                        If zlCommFun.ActualLen(.TextMatrix(i, DI_��ע)) > 200 Then
                            .Row = i: .Col = DI_��ע
                            If gclsPros.FuncType = f���ѡ�� Then
                                Call ShowMessage(vsTmp, """" & .TextMatrix(i, DI_�������) & """�ı�ע����̫����ֻ����200���ַ���100�����֡�")
                                Exit Function
                            Else
                                Call AddErrInfo("""" & .TextMatrix(i, DI_�������) & """�ı�ע����̫����ֻ����200���ַ���100�����֡�", 0, vsTmp)
                            End If
                        End If
                        If Val(.TextMatrix(i, DI_��Ϸ���)) = DT_��Ժ���ZY Then
                            If .TextMatrix(i, DI_��Ժ����) = "" And DiagCellEditable(vsTmp, i, DI_��Ժ����) Then
                                .Row = i: .Col = DI_��Ժ����
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "����д��Ժ���顣")
                                    Exit Function
                                Else
                                    Call AddErrInfo("����д��Ժ���顣", 0, vsTmp)
                                End If
                            End If
                            If .TextMatrix(i, DI_��Ժ���) = "" Then
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "����д��Ժ��ϵĳ�Ժ�����")
                                    Exit Function
                                Else
                                    Call AddErrInfo("����д��Ժ��ϵĳ�Ժ�����", 0, vsTmp)
                                End If
                            End If
                            
                            If Val(.TextMatrix(i - 1, DI_��Ϸ���)) = DT_��Ժ���ZY And InStr(.TextMatrix(FindDiagRow(DT_��Ժ���ZY), DI_��Ժ���), "����") = 0 And InStr(.TextMatrix(i, DI_��Ժ���), "����") > 0 Then
                                .Row = i: .Col = DI_��Ժ���
                                If gclsPros.FuncType = f���ѡ�� Then
                                    Call ShowMessage(vsTmp, "��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������")
                                    Exit Function
                                Else
                                    Call AddErrInfo("��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������", 0, vsTmp)
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_����ID)) <> 0 Then gclsPros.DiseaseIDs = gclsPros.DiseaseIDs & "," & Val(.TextMatrix(i, DI_����ID))
                    If Val(.TextMatrix(i, DI_���ID)) <> 0 Then gclsPros.DiagIDs = gclsPros.DiagIDs & "," & Val(.TextMatrix(i, DI_���ID))
                    '�Ƿ�������Ҫ����������
                    If gclsPros.PatiType = PF_���� Then
                        gclsPros.IsDiagInput = True
                    Else
                        If InStr("," & gclsPros.MustDiagType & ",", "," & Val(.TextMatrix(i, DI_��Ϸ���)) & ",") > 0 Then
                            gclsPros.IsDiagInput = True
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_����)) <> 0 Then
                        gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs <> "", ",", "") & .RowData(i)
                        strTmp = IIf(Trim(.TextMatrix(i, DI_��ϱ���)) = "", "", "(" & .TextMatrix(i, DI_��ϱ���) & ")") & .TextMatrix(i, DI_�������) & IIf(.TextMatrix(i, DI_��ҽ֤��) <> "", "(" & .TextMatrix(i, DI_��ҽ֤��) & ")", "") & IIf(.TextMatrix(i, DI_��ҽ֤��) <> "", "(" & .TextMatrix(i, DI_��ҽ֤��) & ")", "")
                        gclsPros.DiagNames = gclsPros.DiagNames & IIf(gclsPros.DiagNames <> "", ",", "") & strTmp
                        .Cell(flexcpData, i, DI_����) = ""
                        blnHaveSel = True
                    End If
                End If
            Next
        End With
    End If
    If gclsPros.FuncType = f������ҳ And Not blnHaveDaig Then
        If gclsPros.FuncType = f���ѡ�� Then
            Call ShowMessage(gclsPros.CurrentForm.vsDiagXY, "��ҽ��Ϻ���ҽ��϶�û������,����!")
            Exit Function
        Else
            Call AddErrInfo("��ҽ��Ϻ���ҽ��϶�û������,����!", 0, gclsPros.CurrentForm.vsDiagXY)
        End If
    End If
    If gclsPros.DiseaseIDs <> "" Then gclsPros.DiseaseIDs = Mid(gclsPros.DiseaseIDs, 2)
    If gclsPros.DiagIDs <> "" Then gclsPros.DiagIDs = Mid(gclsPros.DiagIDs, 2)
    CheckDiagData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Check����(Optional blnCheck As Boolean) As Boolean
    Dim objCbo As ComboBox, objMSK As MaskEdBox
    Dim curDate As Date
    Dim blnJudge As Boolean
    Dim colErr As New Collection
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    Call ClearErrCol
    Set colErrTmp = colErr  '��ռ���
    curDate = zlDatabase.Currentdate
    With gclsPros.CurrentForm
        If .txtSpecificInfo(SLC_����).Enabled And Not .txtSpecificInfo(SLC_����).Locked Then
            '��Ŀ���볤�ȼ��
            If .txtSpecificInfo(SLC_����).MaxLength <> 0 And .txtSpecificInfo(SLC_����).Text <> "" Then
                strTmp = .txtSpecificInfo(SLC_����).Text
                strTmp = strTmp & .cboSpecificInfo(SLC_����).Text
                If zlCommFun.ActualLen(strTmp) > .txtSpecificInfo(SLC_����).MaxLength Then
                    Call AddErrInfo("�������ݹ�����(����Ŀ������� " & .txtSpecificInfo(SLC_����).MaxLength & " ���ַ��� " & .txtSpecificInfo(SLC_����).MaxLength / 2 & " ������)", 0, .txtSpecificInfo(SLC_����))
                End If
            End If
        End If
        If .txtInfo(GC_����).Enabled And Not .txtInfo(GC_����).Locked Then
            If .txtInfo(GC_����).Text = "" Then
                Call AddErrInfo("���˵���������Ϊ�գ����䲡�˵�������", 0, .txtInfo(GC_����))
            End If
        End If
        If .txtSpecificInfo(SLC_����).Text = "" Then
            Call AddErrInfo("���˵����䲻��Ϊ�գ������벡�˵����䡣", 0, .txtSpecificInfo(SLC_����))
        Else
            'СʱΪ��λ�����ܴ���30�켴720Сʱ
            '����Ϊ��λ�����ܴ���24Сʱ��1440����
            If .cboSpecificInfo(SLC_����).Visible Then
                strTmp = .cboSpecificInfo(SLC_����).Text
                i = decode(strTmp, "��", 200, "��", 2400, "��", 73000, "Сʱ", 720, "����", 1440, 0)
                If Val(.txtSpecificInfo(SLC_����).Text) > i And strTmp <> "" Then
                    Call AddErrInfo("����ֵ�����������" & i & strTmp & IIf(i = 1440 Or i = 720, "����ʹ�ú��ʵ����䵥λ��", "��"), 0, .txtSpecificInfo(SLC_����), .cboSpecificInfo(SLC_����))
                ElseIf Val(.txtSpecificInfo(SLC_����).Text) < 0 Then
                    Call AddErrInfo("����ֵ����Ϊ������", 0, .txtSpecificInfo(SLC_����), .cboSpecificInfo(SLC_����))
                End If
            End If
        End If
        strTmp = ""
        
        '����Ҫ��������ݼ��
        If gclsPros.PatiType = PF_���� Then
            If InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") > 0 Then
                If .cboBaseInfo(BCC_���ʽ).Enabled And Not .cboBaseInfo(BCC_���ʽ).Locked Then
                    strTmp = "���ʽ"
                    If .cboBaseInfo(BCC_���ʽ).ListIndex = -1 Then
                       Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, .cboBaseInfo(BCC_���ʽ))
                    End If
                 End If
            End If
        Else
            If .cboBaseInfo(BCC_���ʽ).Enabled And Not .cboBaseInfo(BCC_���ʽ).Locked Then
                strTmp = "���ʽ"
                If .cboBaseInfo(BCC_���ʽ).ListIndex = -1 Then
                    Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, .cboBaseInfo(BCC_���ʽ))
                End If
            End If
            If .cboBaseInfo(BCC_����).Enabled And Not .cboBaseInfo(BCC_����).Locked Then
                strTmp = "����"
                If .cboBaseInfo(BCC_����).ListIndex = -1 Then
                     Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, .cboBaseInfo(BCC_����))
                End If
            End If
            If .cboBaseInfo(BCC_�Ա�).Enabled And Not .cboBaseInfo(BCC_�Ա�).Locked Then
                strTmp = "BCC_�Ա�"
                If .cboBaseInfo(BCC_�Ա�).ListIndex = -1 Then
                    Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, .cboBaseInfo(BCC_�Ա�))
                End If
            End If
        End If
        If .mskDateInfo(DC_��������).Enabled Then
            blnJudge = .mskDateInfo(DC_��������).Text = Replace(.mskDateInfo(DC_��������).Mask, "#", "_")
            If blnJudge Then
                Call AddErrInfo("�����벡�˵ĳ������ڡ�", 0, .mskDateInfo(DC_��������))
            End If
            If Not IsDate(.mskDateInfo(DC_��������).Text) Then
                Call AddErrInfo("�������ڲ�����Ч�����ڸ�ʽ��", 0, .mskDateInfo(DC_��������))
            End If
            If Format(.mskDateInfo(DC_��������).Text, "yyyy-MM-dd hh:mm") > Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") Then
                Call AddErrInfo("������������Ժʱ��֮��", 0, .mskDateInfo(DC_��������), .mskDateInfo(DC_��Ժʱ��))
            End If
        End If
        Call CheckDiagData(curDate)
    End With
    If gColErr.Count > 0 Or gColWarn.Count > 0 Then
        Call LoadVsErrData
        If Not blnCheck Then
            If gColErr.Count = 0 And gColWarn.Count > 0 Then
                If MsgBox("����" & CStr(gColWarn.Count) & "�����棬�Ƿ����ȫ�����棬����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    Call ClearErrCol
                    Check���� = True
                End If
            ElseIf gColErr.Count > 0 Then
                Check���� = False
                Exit Function
            End If
        End If
    End If
    If Not CheckMedPageChange Then
        gclsPros.InfosChange = False
        gclsPros.IsCheckData = False
        Exit Function
    Else
        gclsPros.IsCheckData = False
    End If
    Check���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CheckMedPageData(ByRef blnDiagnose As Boolean, Optional blnCheck As Boolean) As Boolean
'���ܣ������ҳ�������ݺϷ���
'���أ�blnDiagnose=�Ƿ���д�����
'������
    Dim objTextBox As TextBox, objCbo As ComboBox, objMSK As MaskEdBox, objChk As CheckBox, vsTmp As VSFlexGrid
    Dim blnJudge As Boolean, strTmp As String
    Dim curDate As Date
    Dim lngSize As Long
    Dim str����IDs As String, str���IDs As String
    Dim i As Long, j As Long
    Dim strSql As String, rsTmp As Recordset
    Dim sgnTmp As Single, blnDo As Boolean, blnDoEx As Boolean
    Dim strBirthday As String, strAge As String, strSex As String, strErrIfno As String, str���� As String
    Dim objTmp As Object
    Dim blnDateIsNull As Boolean
    Dim blnBaseInfo As Boolean, strBaseInfo As String
    Dim strMask As String, arrTmp As Variant
    Dim str��סʱ�� As String, strת��ʱ�� As String
    Dim strMsg As String
    Dim colErr As New Collection
    

    blnDiagnose = False
    gclsPros.IsCheckData = True
'    gclsPros.InfosChange = False
    '�����֮ǰ�ľ������
    Call ClearErrCol
    Set colErrTmp = colErr  '��ռ���
    '������鲡���Ƿ��Ŀ����ҳ��������״̬
    If gclsPros.PatiType <> PF_���� And gclsPros.FuncType <> f������ҳ Then
        If Not CheckMecRed(gclsPros.����ID, gclsPros.��ҳID, gclsPros.CurrentForm.Caption, "�޸���ҳ") Then Exit Function
    End If

    curDate = zlDatabase.Currentdate
    With gclsPros.CurrentForm
        'txtInfo�ؼ���ؼ��
        For Each objTextBox In .txtInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                '��Ŀ���볤�ȼ��
                If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                    If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                        Call AddErrInfo("�������ݹ�����(����Ŀ������� " & objTextBox.MaxLength & " ���ַ��� " & objTextBox.MaxLength \ 2 & " ������)", 0, objTextBox)
                    End If
                End If
                Select Case objTextBox.Index
                    Case GC_�໤�����֤��
                        strTmp = objTextBox.Text
                        If strTmp <> "" Then
                            If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                                If zlCommFun.ActualLen(strTmp) = Len(strTmp) Then
                                    If gclsPros.IsMaskID Then strTmp = objTextBox.Tag
                                        '��ʼ��������Ϣ�ӿ�
                                        If gobjPatient Is Nothing Then
                                            On Error Resume Next
                                            Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                                            Err.Clear: On Error GoTo errH
                                            Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
                                        End If
                                        If gobjPatient Is Nothing Then
                                            Call AddErrInfo("����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ����ܽ��в������֤��Ϣ��飡�Ƿ������", 1, objCbo)
                                        End If
                                        If Not gobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno, CDate(gclsPros.InTime)) Then '���֤�Ϸ������Ƿ�ƥ��
                                            '���֤���Ϸ����˳�
                                            Call AddErrInfo(strErrIfno, 0, objTextBox)
                                        End If
                                ElseIf zlCommFun.ActualLen(strTmp) > 18 Then
                                    Call AddErrInfo("���֤�Ų��ܳ���9�����ֻ�18��Ӣ���ַ��ĳ��ȣ����顣", 0, objTextBox)
                                End If
                            End If
                        End If
                End Select
                '����Ҫ��������ݼ��
                If objTextBox.Index = GC_Email Then
                     If objTextBox.Text <> "" Then
                        If InStr(objTextBox.Text, "@") <= 1 Or InStr(objTextBox.Text, ".") <= 3 Or InStr(objTextBox.Text, "@") > InStr(objTextBox.Text, ".") Then
                            Call AddErrInfo("�����Email�ĸ�ʽ����ȷ����ȷ��ʽ��""XXX@XX.XX""��", 0, objTextBox)
                        End If
                    End If
                ElseIf objTextBox.Index = GC_ת��2 Then
                     If objTextBox.Text <> "" Then
                        If .txtInfo(GC_ת��1).Text = "" Then
                            Call AddErrInfo("û����������ת�ƿ��ң����������롣", 0, .txtInfo(GC_ת��1), objTextBox)
                        ElseIf .txtInfo(GC_ת��1).Text = objTextBox.Text Or .txtInfo(GC_ת��3).Text = objTextBox.Text Then
                            Call AddErrInfo("ת�Ƶ��������Ҳ�Ӧ����ͬ��", 0, objTextBox, .txtInfo(IIf(.txtInfo(GC_ת��1).Text = objTextBox.Text, GC_ת��1, GC_ת��3)))
                        End If
                    Else
                         If .txtInfo(GC_ת��3).Text <> "" Then
                            Call AddErrInfo("û����������ת�ƿ��ң����������롣", 0, objTextBox, .txtInfo(GC_ת��3))
                        End If
                    End If
                ElseIf objTextBox.Index = GC_31������סԺ Then
                    If Trim(objTextBox.Text) = "" Then
                        Call AddErrInfo(.cboBaseInfo(BCC_����Ժ�ƻ�����).Text & "��Ŀ��û����д��", 0, objTextBox)
                    End If
                ElseIf objTextBox.Index = BCC_�����ڼ� Then
                    If objTextBox.Text <> "" Then
                        If .cboBaseInfo(BCC_�����ڼ�).Text = "" And objTextBox.Index = GC_����ԭ�� Then
                            Call AddErrInfo("����������ԭ�򣬵�û��¼�������ڼ䣬�Ƿ������", 1, .cboBaseInfo(BCC_�����ڼ�), objTextBox)
                        End If
                    End If
                ElseIf gclsPros.FuncType = f������ҳ Then
                    If objTextBox.Text = "" Then
                        If objTextBox.Index = GC_���� Then
                            Call AddErrInfo("���˵���������Ϊ�գ����䲡�˵�������", 0, objTextBox)
                        ElseIf objTextBox.Index = GC_��Ժ���� Then
                            Call AddErrInfo("���˵���Ժ���Ҳ���Ϊ�գ����䲡�˵���Ժ���ҡ�", 0, objTextBox)
                        ElseIf objTextBox.Index = GC_��Ժ���� Then
                            Call AddErrInfo("���˵ĳ�Ժ���Ҳ���Ϊ�գ����䲡�˵ĳ�Ժ���ҡ�", 0, objTextBox)
                        End If
                    End If
                End If
            End If
        Next

        '��ȡ�ϴγ�Ժʱ���Լ��´���Ժʱ��
        If gclsPros.FuncType = f������ҳ Then
            If gclsPros.InTime = "" Then
                Call AddErrInfo("���˵���Ժ���ڲ���Ϊ�գ������벡�˵���Ժ���ڡ�", 0, .mskDateInfo(DC_��Ժʱ��))
            End If

            If gclsPros.OutTime = "" Then
                Call AddErrInfo("���˵ĳ�Ժ���ڲ���Ϊ�գ������벡�˵ĳ�Ժ���ڡ�", 0, .mskDateInfo(DC_��Ժʱ��))
            End If
            If gclsPros.InTime > Format(curDate, "yyyy-MM-dd hh:mm:ss") Then
                Call AddErrInfo("���˵���Ժ���ڲ��������ڵ�ǰʱ�䡣", 0, .mskDateInfo(DC_��Ժʱ��))
            End If
            If gclsPros.OutTime > Format(curDate, "yyyy-MM-dd hh:mm:ss") Then
                Call AddErrInfo("���˵ĳ�Ժ���ڲ��������ڵ�ǰʱ�䡣", 0, .mskDateInfo(DC_��Ժʱ��))
            End If
            If gclsPros.InTime > gclsPros.OutTime Then
                Call AddErrInfo("���˵ĳ�Ժ���ڲ���������Ժʱ�䡣", 0, .mskDateInfo(DC_��Ժʱ��), .mskDateInfo(DC_��Ժʱ��))
            End If

            strSql = "Select ��ҳid, �´���Ժ, �ϴγ�Ժ" & vbNewLine & _
                    "From (Select ��ҳid, Lead(��Ժ����, 1, Null) Over(Order By ��ҳid) �´���Ժ, Lag(��Ժ����, 1, Null) Over(Order By ��ҳid) �ϴγ�Ժ" & vbNewLine & _
                    "       From ������ҳ" & vbNewLine & _
                    "       Where ����id = [1])" & vbNewLine & _
                    "Where ��ҳid = [2]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ٽ���ҳID�����Ժʱ��", gclsPros.����ID, gclsPros.��ҳID)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!�ϴγ�Ժ) Then
                    strTmp = Format(rsTmp!�ϴγ�Ժ & "", "yyyy-MM-dd hh:mm")
                    '�������Ժʱ����
                    If Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") < strTmp Then
                        Call AddErrInfo("������Ժ�����������ϴεĳ�Ժ����(" & strTmp & ")��", 0, .mskDateInfo(DC_��Ժʱ��))
                    End If
                End If
                If Not IsNull(rsTmp!�´���Ժ) Then
                    strTmp = Format(rsTmp!�´���Ժ & "", "yyyy-MM-dd hh:mm")
                    If Format(gclsPros.OutTime, "yyyy-MM-dd hh:mm") > strTmp Then
                        Call AddErrInfo("���γ�Ժ������������һ�ε���Ժ����(" & strTmp & ")��", 0, .mskDateInfo(DC_��Ժʱ��))
                    End If
                End If
            End If
        End If

        'txtSpecificInfo�ؼ���ؼ��
        For Each objTextBox In .txtSpecificInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                '��Ŀ���볤�ȼ��
                If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                    strTmp = objTextBox.Text
                    If objTextBox.Index = SLC_���� Then
                        strTmp = strTmp & .cboSpecificInfo(SLC_����).Text
                    End If
                    If zlCommFun.ActualLen(strTmp) > objTextBox.MaxLength Then
                        Call AddErrInfo("�������ݹ�����(����Ŀ������� " & objTextBox.MaxLength & " ���ַ��� " & objTextBox.MaxLength / 2 & " ������)", 0, objTextBox)
                    End If
                End If

                Select Case objTextBox.Index
                    Case SLC_��ͥ�绰, SLC_��λ�绰, SLC_��ϵ�˵绰
                        strMask = "1234567890-()"
                        For i = 1 To Len(objTextBox.Text)
                            If InStr(strMask, Mid(objTextBox.Text, i, 1)) = 0 Then
                                Call AddErrInfo("��������ݳ����˷Ƿ��ַ�����������������ַ�Ϊ(" & strMask & ")", 0, objTextBox)
                                Exit For
                            End If
                        Next
                    Case SLC_סԺ��, SLC_�����ʱ�, SLC_��ͥ�ʱ�, SLC_��λ�ʱ�, SLC_���ȴ���, SLC_�ɹ�����, _
                        SLC_��������, SLC_����ʱ����Ժǰ_Сʱ, SLC_����ʱ����Ժǰ_����, SLC_����ʱ����Ժ��_����, SLC_����ʱ����Ժ��_Сʱ, _
                        SLC_����ʱ����Ժǰ_��, SLC_����ʱ����Ժ��_��, SLC_������ʹ��, SLC_��֢�໤��, SLC_��֢�໤Сʱ, SLC_QQ, SLC_��Ժ����, SLC_Ժ�ڻ���
                        strMask = "1234567890"
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("������������󣬸���ֻ�ܹ�����������", 0, objTextBox)
                            End If
                        End If
                    Case SLC_���ϸ��, SLC_��ѪС��, SLC_��Ѫ��, SLC_��ȫѪ, SLC_��׵���, SLC_�������, SLC_ICU, _
                         SLC_CCU, SLC_һ������, SLC_��������, SLC_��������, SLC_�ػ�, SLC_Լ����ʱ��, SLC_���, SLC_����, SLC_���ϴ�סԺʱ��
                        strMask = "1234567890."
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("������������󣬸���ֻ�ܹ��������֡�", 0, objTextBox)
                            End If
                        End If
                    Case SLC_��������������, SLC_��������Ժ����
                        If objTextBox.Text <> "" Then
                            If InStr(objTextBox.Text, ";") > 0 Then
                                arrTmp = Split(objTextBox.Text, ";")
                                For i = LBound(arrTmp) To UBound(arrTmp)
                                    If Not IsNumeric(arrTmp(i)) Then
                                        Call AddErrInfo("�������������,����ֻ�ܹ������֣����ж�����������������������������á�;���ָ�����", 0, objTextBox)
                                        Exit For
                                    End If
                                Next
                            Else
                                 If Not IsNumeric(objTextBox.Text) Then
                                    Call AddErrInfo("�������������,����ֻ�ܹ������֣����ж�����������������������������á�;���ָ�����", 0, objTextBox)
                                End If
                            End If
                        End If
                    Case SLC_Apgar
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("������������󣬸���ֻ�ܹ�����������", 0, objTextBox)
                            ElseIf Val(objTextBox.Text) > 10 Then
                                Call AddErrInfo("�����ֵֻ����0-10 ֮�䡣", 0, objTextBox)
                            End If
                        End If
                End Select

                '����Ҫ��������ݼ��
                Select Case objTextBox.Index
                    Case SLC_סԺ��
                        If objTextBox.Text = "" Then
                            Call AddErrInfo("���˵�סԺ�Ų���Ϊ�գ������벡�˵�סԺ�š�", 0, objTextBox)
                        Else
                            If gclsPros.InNo <> "0" Then
                                If Trim(objTextBox.Text) <> gclsPros.InNo Then
                                    Call AddErrInfo("�ò��˵�סԺ���ѷ����ı䣬���顣", 0, objTextBox)
                                End If
                            End If
                        End If
                    Case SLC_����
                        If objTextBox.Text = "" Then
                            Call AddErrInfo("���˵����䲻��Ϊ�գ������벡�˵����䡣", 0, objTextBox)
                        Else
                            'СʱΪ��λ�����ܴ���30�켴720Сʱ
                            '����Ϊ��λ�����ܴ���24Сʱ��1440����
                            If .cboSpecificInfo(SLC_����).Visible Then
                                strTmp = .cboSpecificInfo(SLC_����).Text
                                i = decode(strTmp, "��", 200, "��", 2400, "��", 73000, "Сʱ", 720, "����", 1440, 0)
                                If Val(objTextBox.Text) > i And strTmp <> "" Then
                                    Call AddErrInfo("����ֵ�����������" & i & strTmp & IIf(i = 1440 Or i = 720, "����ʹ�ú��ʵ����䵥λ��", "��"), 0, objTextBox, .cboSpecificInfo(SLC_����))
                                ElseIf Val(objTextBox.Text) < 0 Then
                                    Call AddErrInfo("����ֵ����Ϊ������", 0, objTextBox, .cboSpecificInfo(SLC_����))
                                End If
                            End If
                        End If
                    Case SLC_Ӥ�׶�����
                        If objTextBox <> "" Then
                            'СʱΪ��λ�����ܴ���30�켴720Сʱ
                            '����Ϊ��λ�����ܴ���24Сʱ��1440����
                            If .cboSpecificInfo(SLC_Ӥ�׶�����).Visible Then
                                strTmp = .cboSpecificInfo(SLC_Ӥ�׶�����).Text
                                i = decode(strTmp, "��", 12, "��", 365, "Сʱ", 720, "����", 1440, 0)
                                If Val(objTextBox.Text) > i Then
                                    Call AddErrInfo("Ӥ������ֵ�����������" & i & strTmp & IIf(i = 1440 Or i = 720, "����ʹ�ú��ʵ����䵥λ��", "��"), 0, objTextBox, .cboSpecificInfo(SLC_Ӥ�׶�����))
                                ElseIf Val(objTextBox.Text) < 0 Then
                                    Call AddErrInfo("Ӥ������ֵ����Ϊ������", 0, objTextBox, .cboSpecificInfo(SLC_Ӥ�׶�����))
                                End If
                            End If
                        End If
                    Case SLC_Ӥ�׶�����_DAY
                        If objTextBox.Visible Then
                            '��Ϊ��λ�����ܴ���30��Ҳ����С��0
                            strTmp = Trim(objTextBox.Text)
                            If strTmp = "" Then
                                If Trim(.txtSpecificInfo(SLC_Ӥ�׶�����).Text) <> "" Then
                                    Call AddErrInfo("Ӥ�����䲻��1���µ�����������Ϊ�ա�", 0, objTextBox, .cboSpecificInfo(SLC_Ӥ�׶�����))
                                End If
                            Else
                                If Trim(.txtSpecificInfo(SLC_Ӥ�׶�����).Text) = "" Then
                                    Call AddErrInfo("Ӥ����������䲻����Ϊ�ա�", 0, .txtSpecificInfo(SLC_Ӥ�׶�����), .cboSpecificInfo(SLC_Ӥ�׶�����))
                                ElseIf strTmp Like "0*" And Len(strTmp) > 1 Then
                                    Call AddErrInfo("Ӥ�����䲻��1���µ�������������Ч��������", 0, objTextBox, .cboSpecificInfo(SLC_Ӥ�׶�����))
                                ElseIf Val(strTmp) >= 30 Or Val(strTmp) < 0 Then
                                    Call AddErrInfo("Ӥ�����䲻��1���µ�����������ȡֵ��ΧΪ���ڵ���0��С��30��������", 0, objTextBox, .cboSpecificInfo(SLC_Ӥ�׶�����))
                                End If
                            End If
                        End If
                    Case SLC_���ȴ���
                        '�ɹ��������ܳ������ȴ���,���˳�Ժ���Ϊ������ʱ�򣬳ɹ��������Ե������ȴ�������Ϊ���ܲ���û�����Ⱦ�����
                        If Val(.txtSpecificInfo(SLC_�ɹ�����).Text) > Val(objTextBox.Text) Then
                            Call AddErrInfo("�ɹ��������ܳ������ȴ�����", 0, .txtSpecificInfo(SLC_�ɹ�����), .txtSpecificInfo(SLC_���ȴ���))
                        End If

                        If objTextBox.Text <> "" Then
                            strTmp = .vsDiagXY.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_��Ժ���)
                        End If
                    Case SLC_��������
                        If Val(objTextBox.Text) <= 0 Then
                            Call AddErrInfo("��������ȷ���������ޡ�", 0, objTextBox)
                        End If
                    Case SLC_סԺ����
                        i = DateDiff("d", CDate(gclsPros.InTime), CDate(gclsPros.OutTime))
                        If i = 0 Then i = 1
                        If i <> Val(objTextBox.Text) Then
                            Call AddErrInfo("���˵�סԺ��������ȷ�������Ժʱ���Ƿ���ȷ��", 0, objTextBox, .mskDateInfo(DC_��Ժʱ��))
                        End If
                End Select
            End If
        Next
        'txtAdressInfo�ؼ���ؼ��
        On Error Resume Next
        For Each objTextBox In .txtAdressInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                strTmp = decode(objTextBox.Index, ADRC_�����ص�, "�����ص�", ADRC_����, "����", ADRC_��סַ, "��סַ", ADRC_���ڵ�ַ, "���ڵ�ַ", ADRC_��ϵ�˵�ַ, "��ϵ�˵�ַ", ADRC_��������, "����")
                '��Ŀ���볤�ȼ��
                blnJudge = .padrInfo(objTextBox.Index).MaxLength '�жϿؼ��Ƿ����
                blnJudge = Err.Number = 0: Err.Clear
                If blnJudge And gclsPros.IsStructAdress Then    '��Ҫ����ַ�ؼ�������
                    If .padrInfo(objTextBox.Index).CheckNullValue() <> "" Then
                        Call AddErrInfo(strTmp & "��" & .padrInfo(objTextBox.Index).CheckNullValue() & "��δ���룬���顣", 0, .padrInfo(objTextBox.Index))
                    End If
                    If .padrInfo(objTextBox.Index).MaxLength > 0 Then
                        If zlCommFun.ActualLen(.padrInfo(objTextBox.Index).Value) > .padrInfo(objTextBox.Index).MaxLength Then
                            Call AddErrInfo(strTmp & "������̫�������顣(����Ŀ������� " & .padrInfo(objTextBox.Index).MaxLength & " ���ַ��� " & .padrInfo(objTextBox.Index).MaxLength \ 2 & " ������)", 0, .padrInfo(objTextBox.Index))
                        End If
                    End If
                Else '��Ҫ���TextBox������
                    If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                        If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                            Call AddErrInfo(strTmp & "�����ݹ��������顣(����Ŀ������� " & objTextBox.MaxLength & " ���ַ��� " & objTextBox.MaxLength \ 2 & " ������)", 0, objTextBox)
                        End If
                    End If
                     '����Ҫ��������ݼ��
                    If objTextBox.Index = ADRC_�������� Then
                        If objTextBox.Text = "" Then
                            If gclsPros.FuncType = f������ҳ Then
                                Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, objTextBox)
                            Else
                                If gclsPros.Check���� = 1 Then
                                    Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, objTextBox)
                                ElseIf gclsPros.Check���� = 2 Then
                                    Call AddErrInfo("û�����벡�˵�" & strTmp & ",�Ƿ������", 1, objTextBox)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        strTmp = ""
        'cboBaseInfo���
        For Each objCbo In .cboBaseInfo
            strTmp = ""
            If objCbo.Enabled And Not objCbo.Locked Then
                '����Ҫ��������ݼ��
                If gclsPros.PatiType = PF_���� Then
                    If InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") > 0 And objCbo.Index = BCC_���ʽ Then strTmp = "���ʽ"
                Else
                    strTmp = decode(objCbo.Index, BCC_���ʽ, "���ʽ", BCC_����, "����", BCC_����, "����", BCC_ְҵ, "ְҵ", BCC_��Ժ���, "��Ժ����", BCC_�Ա�, "�Ա�", "")
                End If
                If strTmp <> "" Then
                    If objCbo.ListIndex = -1 Then
                        Call AddErrInfo("�����벡�˵�" & strTmp & "��", 0, objCbo)
                    End If
                End If
                '�������ݵ���Ч�Լ��
                Select Case objCbo.Index
                    Case BCC_Ѫ��
                        If objCbo.Text = "" Then
                            Call AddErrInfo("������д�ò��˵�Ѫ��.", 0, objCbo)
                        End If
                    Case BCC_RH
                        If objCbo.Text = "" Then
                            Call AddErrInfo("������д�ò��˵�RH", 0, objCbo)
                        End If
                    Case BCC_���� '15������ӦΪδ��
                        If objCbo.Text <> "" And objCbo.ListIndex <> -1 Then
                            If InStr(objCbo.Text, "δ��") = 0 And InStr(objCbo.Text, "����") = 0 Then
                                If IsDate(.mskDateInfo(DC_��������).Text) Then
                                    If DateDiff("yyyy", CDate(.mskDateInfo(DC_��������).Text), curDate) < 15 Then
                                        Call AddErrInfo("�ò�������С��15�꣬����״��Ӧ��дΪδ����������Ƿ������", 1, objCbo)
                                    End If
                                End If
                            End If
                        End If
                    Case BCC_��Ժ��� '��Ժ����ΪΣʱ��Ҫ��������
                        If InStr(objCbo.Text, "Σ") > 0 And Val(.txtSpecificInfo(SLC_���ȴ���).Text) = 0 Then
                            Call AddErrInfo("�ò�����Ժ����ΪΣ����û�н������ȣ��Ƿ������", 1, .txtSpecificInfo(SLC_���ȴ���), objCbo)
                        End If
                    Case BCC_��ǰ������ '��д����ǰ�����󣬱�����д�������
                        If .vsOPS.TextMatrix(1, PI_��������) = "" And objCbo.ListIndex > 0 Then
                            Call AddErrInfo("û����д�������,��ǰ������ֻ��ѡ��""δ��""��", 0, objCbo)
                        End If
                    Case BCC_��Ժ��ʽ '�����Ժ��ʽ�������������Ժ����Ƿ�Ϊ����
                        If InStr(objCbo.Text, "����") > 0 Then
                            strTmp = .vsDiagXY.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_��Ժ���)
                            If strTmp = "" And gclsPros.IsTCM Then
                                strTmp = .vsDiagZY.TextMatrix(FindDiagRow(DT_��Ժ���ZY), DI_��Ժ���)
                            End If
                            If strTmp <> "" Then
                                If InStr(strTmp, "����") = 0 Then
                                    Call AddErrInfo("������������Ϊ��������Ժ��ʽΪ������", 0, objCbo)
                                End If
                            End If
                        End If
                    Case BCC_���֤
                        If objCbo.Enabled And Not objCbo.Locked Then
                            '�����֤�Ž�����֤
                            If objCbo.Index = BCC_���֤ Then
                                strTmp = objCbo.Text
                                If strTmp <> "" Then
                                    If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                                        If zlCommFun.ActualLen(strTmp) = Len(strTmp) Then
                                            If gclsPros.IsMaskID Then strTmp = objCbo.Tag
                                                '��ʼ��������Ϣ�ӿ�
                                                If gobjPatient Is Nothing Then
                                                    On Error Resume Next
                                                    Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                                                    Err.Clear: On Error GoTo errH
                                                    Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
                                                End If
                                                If gobjPatient Is Nothing Then
                                                    Call AddErrInfo("����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ����ܽ��в������֤��Ϣ��飡�Ƿ������", 1, objCbo)
                                                End If
                                                If gobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno, CDate(gclsPros.InTime)) Then '���֤�Ϸ������Ƿ�ƥ��
                                                    
                                                                                                        If Val(zlDatabase.GetPara(279, 100)) = 1 Then
                                                        strSql = "select 1 from ������Ϣ a where a.���֤��=[1] and a.����id<>[2] and rownum<2"
                                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, strTmp, gclsPros.����ID)
                                                        If Not rsTmp.EOF Then Call AddErrInfo("�����֤���Ѿ�������ͬһ���ֻ֤�ܶ�Ӧһ����������!", 0, objCbo)
                                                    End If

                                                    strBaseInfo = "": Set objTmp = Nothing
                                                    If Not Trim(.txtSpecificInfo(SLC_����).Text) Like "Լ*" Or Trim(.txtSpecificInfo(SLC_����).Text) = "����" Then
                                                        If Format(strBirthday, "yyyy-MM-dd") <> Format(.mskDateInfo(DC_��������).Text, "yyyy-MM-dd") Then
                                                            strBaseInfo = "��������"
                                                            Set objTmp = objCbo
                                                        End If
                                                                If Format(.mskDateInfo(DC_��������).Text, "HH:MM") <> "00:00" Then
                                                                        strBirthday = strBirthday & " " & Format(.mskDateInfo(DC_��������).Text, "HH:MM")
                                                                End If
                                                        If gclsPros.Sex <> strSex Then
                                                            strBaseInfo = strBaseInfo & IIf(strBaseInfo <> "", "��", "") & "�Ա�"
                                                            Set objTmp = .cboBaseInfo(BCC_�Ա�)
                                                        End If
                                                        str���� = .txtSpecificInfo(SLC_����).Text & IIf(.cboSpecificInfo(SLC_����).Visible, .cboSpecificInfo(SLC_����).Text, "")
                                                        If strAge <> str���� Then
                                                            strBaseInfo = strBaseInfo & IIf(strBaseInfo <> "", "��", "") & "����"
                                                            Set objTmp = .txtSpecificInfo(SLC_����)
                                                                        If Trim(str����) Like "*Сʱ*����" Or Trim(str����) Like "*����" Or Trim(str����) Like "*��*Сʱ" Or Trim(str����) Like "*Сʱ" Then
                                                                        strAge = .txtSpecificInfo(SLC_����).Text & IIf(.cboSpecificInfo(SLC_����).Visible, .cboSpecificInfo(SLC_����).Text, "")
                                                                        End If
                                                        End If
                                                    End If
                                                    If strBaseInfo <> "" Then
                                                        If InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") = 0 Or gclsPros.FuncType = f������ҳ Then
                                                            Call AddErrInfo("���֤�����ȡ��" & strBaseInfo & "�뵱ǰ�����" & strBaseInfo & "��������Ƿ������", 1, objTmp)
                                                        Else
                                                            Call AddErrInfo("���֤�����ȡ��" & strBaseInfo & "�뵱ǰ�����" & strBaseInfo & "��������Ƿ�������������Զ����½����ϵ�" & strBaseInfo & "��", 1, objTmp)
                                                            blnBaseInfo = True
                                                        End If
                                                    End If
                                            Else '���֤���Ϸ����˳�
                                                Call AddErrInfo(strErrIfno, 0, objCbo)
                                            End If
                                        ElseIf zlCommFun.ActualLen(strTmp) > 18 Then
                                            Call AddErrInfo("���֤�Ų��ܳ���9�����ֻ�18��Ӣ���ַ��ĳ��ȣ����顣", 0, objCbo)
                                        End If
                                    End If
                                ElseIf gclsPros.FuncType = f������ҳ Then
                                    Call AddErrInfo("���֤����û�����룬�Ƿ������", 1, objCbo)
                                End If
                            End If
                        End If
                End Select
            End If
        Next
        If gclsPros.PatiType <> PF_���� Then
            'cboManInfo���
            strTmp = ""
            For Each objCbo In .cboManInfo
                If Not objCbo.Locked Then
                    Select Case objCbo.Index
                        Case MC_������
                            If objCbo.Text = "" Then strTmp = strTmp & ";������"
                        Case MC_���λ�����
                            If objCbo.Text = "" Then strTmp = strTmp & ";����ҽʦ"
                        Case MC_����ҽʦ
                            If objCbo.Text = "" Then strTmp = strTmp & ";����ҽʦ"
                        Case MC_סԺҽʦ
                            If objCbo.Text = "" Then strTmp = strTmp & ";סԺҽʦ"
                        Case MC_��ĿԱ
                            If gclsPros.FuncType = f������ҳ Then
                                If objCbo.Text = "" Then
                                    Call AddErrInfo("�������ĿԱ��", 0, objCbo)
                                End If
                            End If
                    End Select
                End If
            Next
            If UBound(Split(strTmp, ";")) = 4 Then '����Ҫ��������ݼ��
                Call AddErrInfo("���ڿ����Ρ�����ҽʦ������ҽʦ��סԺҽʦ֮������ѡ��һλ��", 0, .cboManInfo(MC_������), .cboManInfo(MC_���λ�����), .cboManInfo(MC_����ҽʦ), .cboManInfo(MC_סԺҽʦ))
            End If
            strTmp = ""
        End If
        'mskDateInfo���
        For Each objMSK In .mskDateInfo
            If objMSK.Enabled Then
                blnJudge = objMSK.Text = Replace(objMSK.Mask, "#", "_")
                '����Ҫ��������ݼ��
                Select Case objMSK.Index
                    Case DC_ȷ������
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("ȷ�����ڲ�����Ч�����ڸ�ʽ��", 0, objMSK)
                            End If
                            If gclsPros.FuncType = f������ҳ Then
                                If Not Between(Format(objMSK.Text, "yyyy-MM-dd"), Format(gclsPros.InTime, "yyyy-MM-dd"), _
                                    Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd")) Then
                                    Call AddErrInfo("ȷ�����ڱ�������Ժʱ��ͳ�Ժʱ��֮�䡣", 0, objMSK)
                                End If
                            Else
                                If objMSK.Mask = "####-##-##" Then
                                    If Not Between(Format(objMSK.Text, "yyyy-MM-dd"), Format(gclsPros.InTime, "yyyy-MM-dd"), _
                                        Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd")) Then
                                        Call AddErrInfo("ȷ�����ڱ�������Ժʱ��ͳ�Ժʱ��֮�䡣", 0, objMSK)
                                    End If
                                Else
                                    If Not Between(Format(objMSK.Text, "yyyy-MM-dd hh:mm"), Format(gclsPros.InTime, "yyyy-MM-dd hh:mm"), _
                                        Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd hh:mm")) Then
                                        Call AddErrInfo("ȷ�����ڱ�������Ժʱ��ͳ�Ժʱ��֮�䡣", 0, objMSK)
                                    End If
                                End If
                            End If
                        ElseIf .chkInfo(CHK_�Ƿ�ȷ��).Value = 1 Then
                            Call AddErrInfo("������ȷ�����ڡ�", 0, objMSK)
                        End If
                    Case DC_��������
                        If blnJudge Then
                            Call AddErrInfo("�����벡�˵ĳ������ڡ�", 0, objMSK)
                        End If
                        If Not IsDate(objMSK.Text) Then
                            Call AddErrInfo("�������ڲ�����Ч�����ڸ�ʽ��", 0, objMSK)
                        End If
                        If Format(objMSK.Text, "yyyy-MM-dd hh:mm") > Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") Then
                            Call AddErrInfo("������������Ժʱ��֮��", 0, objMSK, .mskDateInfo(DC_��Ժʱ��))
                        End If
                    Case DC_��������
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("��������ȷ�ķ������ڡ�", 0, objMSK)
                            Else
                                If Not IsDate(.mskDateInfo(DC_����ʱ��).Text) And .mskDateInfo(DC_����ʱ��).Text <> "__:__" Then
                                    Call AddErrInfo("��������ȷ�ķ���ʱ�䡣", 0, .mskDateInfo(DC_����ʱ��))
                                End If
                                strTmp = IIf(IsDate(.mskDateInfo(DC_����ʱ��).Text), " " & .mskDateInfo(DC_����ʱ��).Text, "")
                                If CDate(objMSK.Text & strTmp) >= CDate(Format(curDate, GetFormat(objMSK.Tag) & IIf(strTmp = "", "", " HH:mm"))) Then
                                    Call AddErrInfo("����ʱ��Ӧ�����ڵ�ǰʱ�䡣", 0, objMSK)
                                End If
                            End If
                        End If
                    Case DC_����ʱ��
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("����ʱ�䲻����Ч�����ڸ�ʽ��", 0, objMSK)
                            End If
                            If Format(objMSK.Text, "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                Call AddErrInfo("����ʱ��Ӧ����Ժʱ����", 0, objMSK)
                            End If
                        End If
                    Case DC_��Ŀ����
                        If blnJudge Then
                            Call AddErrInfo("�������Ŀ���ڡ�", 0, objMSK)
                        End If
                        If Not IsDate(objMSK.Text) Then
                            Call AddErrInfo("��Ŀ���ڲ�����Ч�����ڸ�ʽ��", 0, objMSK)
                        End If
                    Case DC_�ʿ�����
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("�ʿ����ڲ�����Ч�����ڸ�ʽ��", 0, objMSK)
                            ElseIf Format(objMSK.Text, "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                                Call AddErrInfo("�ʿ����ڲ���С����Ժ���ڡ�", 0, objMSK)
                            End If
                        End If
                End Select
            End If
        Next
        '��Ժ�����嵥���
        If gclsPros.FuncType = f������ҳ And gclsPros.InputOutList Then
            strSql = "Select A.����, A.����, B.����, B.Id" & vbNewLine & _
                    "From ��Ժ�����嵥 A, ���ű� B" & vbNewLine & _
                    "Where A.����id = B.Id And A.סԺ�� = [1]" & vbNewLine & _
                    "Order By A.���� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, .txtSpecificInfo(SLC_סԺ��).Text)
            If rsTmp.RecordCount = 0 Then
                If gclsPros.OpenMode = EM_�༭ Then
                    Call AddErrInfo("סԺ��Ϊ" & .txtSpecificInfo(SLC_סԺ��).Text & "�Ĳ��˻�û�б�¼��סԺ�ձ��ĳ�Ժ�����嵥�У��Ƿ������", 1, .txtSpecificInfo(SLC_סԺ��))
                Else
                    Call AddErrInfo("סԺ��Ϊ" & .txtSpecificInfo(SLC_סԺ��).Text & "�Ĳ��˻�û�б�¼��סԺ�ձ��ĳ�Ժ�����嵥�С�", 0, .txtSpecificInfo(SLC_סԺ��))
                End If
            Else
                If Not IsNull(rsTmp!����) Then
                    If zlCommFun.TruncateDate(gclsPros.OutTime) <> zlCommFun.TruncateDate(rsTmp!���� & "") Then
                        Call AddErrInfo("��סԺ�ձ��иò��˵ĳ�Ժ����(" & Format(rsTmp!���� & "", "yyyy-MM-dd") & _
                                ")�뵱ǰ��д�Ĳ������Ƿ������", 1, .mskDateInfo(DC_��Ժʱ��))
                    End If
                End If
                If gclsPros.��Ժ����ID <> Val(rsTmp!ID) And gclsPros.��Ժ����ID <> 0 Then
                    Call AddErrInfo("��סԺ�ձ��иò��˵ĳ�Ժ����(" & .txtInfo(GC_��Ժ����).Text & ")�뵱ǰ��д�Ĳ������Ƿ������", 1, .txtInfo(GC_��Ժ����))
                End If
            End If
        End If
    End With

    If Not gclsPros.Is��ʿվ Then
        '���ļ��
        '----------------------------------------------------------------------------------------
        If gclsPros.FuncType = f������ҳ Then
            Set vsTmp = gclsPros.CurrentForm.vsTransfer
            With vsTmp
                For i = .FixedCols To .Cols - 1
                    If .TextMatrix(DR_ת�ƿ���, i) <> "" Then
                        If .TextMatrix(DR_ת�ƿ���, i) = .TextMatrix(DR_ת�ƿ���, i - 1) Then
                            .Row = DR_ת�ƿ���: .Col = i
                            Call AddErrInfo("��" & i & "�����" & i - 1 & "��ת�������ͬ,����!��Ҫ�������������밴Insert����", 0, vsTmp)
                        End If
                    ElseIf i <> .Cols - 1 Then '�����Լ�����ת�ƿ���
                        If .TextMatrix(DR_ת�ƿ���, i + 1) <> "" Then
                            .Row = DR_ת�ƿ���: .Col = i
                            Call AddErrInfo("��" & i & "��û��ת�����,����" & i + 1 & "�д���ת�����,����!��Ҫɾ�������밴Delete����", 0, vsTmp)
                        End If
                    End If
                    If .TextMatrix(DR_ת�ƿ���, i) = "" And Trim(.TextMatrix(DR_ת�ƿ���, i)) <> "" Then
                        .Row = DR_ת�ƿ���: .Col = i
                        Call AddErrInfo("��" & i & "����ת��ʱ�䵫δ��ת�����!��Ҫɾ�������밴Delete����", 0, vsTmp)
                    End If
                Next
            End With
        End If

        gclsPros.Have���� = False
        If gclsPros.PatiType <> PF_���� Then
            Set vsTmp = gclsPros.CurrentForm.vsOPS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, PI_��������)) <> "" Then
                        gclsPros.Have���� = True
                    End If
                Next
            End With
        End If

        Call CheckDiagData(curDate)
        '����ҩ������
        blnJudge = True
        Set vsTmp = gclsPros.CurrentForm.vsAller
        With vsTmp
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, AI_����ҩ��)) <> "" Then
                    If blnJudge Then
                        If gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼).Value = 1 And Trim(.TextMatrix(i, AI_����ҩ��)) <> "��" Then
                            Call AddErrInfo("�ò��˴��ڹ�����¼�����ܹ�ѡ�޹�����¼��", 0, gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼))
                        End If
                    End If
                    blnJudge = False
                    If zlCommFun.ActualLen(.TextMatrix(i, AI_����ҩ��)) > 60 Then
                        .Row = i: .Col = AI_����ҩ��
                        Call AddErrInfo("����ҩ����̫����ֻ����60���ַ���30�����֡�", 0, vsTmp)
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, AI_������Ӧ)) > 100 Then
                        .Row = i: .Col = AI_������Ӧ
                        Call AddErrInfo("������Ӧ����̫����ֻ����100���ַ���50�����֡�", 0, vsTmp)
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, AI_����ҩ��)) <> "" And Format(.TextMatrix(i, AI_����ʱ��), "yyyy-mm-dd") = Format(.TextMatrix(j, AI_����ʱ��), "yyyy-mm-dd") Then
                            blnDateIsNull = False
                            If .TextMatrix(j, AI_����ʱ��) = "" Then blnDateIsNull = True
                            If .TextMatrix(j, AI_����ҩ��) = .TextMatrix(i, AI_����ҩ��) Then
                                .Row = i: .Col = AI_����ҩ��
                                Call AddErrInfo("����" & IIf(blnDateIsNull, "���ڹ���ʱ��Ϊ�յ���ͬ�Ĺ���ҩ���¼��", Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��"), 0, vsTmp)
                            ElseIf Val(.TextMatrix(i, AI_ҩ��ID)) <> 0 And .TextMatrix(i, AI_ҩ��ID) = .TextMatrix(j, AI_ҩ��ID) Then
                                .Row = i: .Col = AI_����ҩ��
                                Call AddErrInfo("����" & IIf(blnDateIsNull, "���ڹ���ʱ��Ϊ�յ���ͬ�Ĺ���ҩ���¼��", Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��"), 0, vsTmp)
                            ElseIf .TextMatrix(i, AI_����Դ����) <> "" And .TextMatrix(i, AI_����Դ����) = .TextMatrix(j, AI_����Դ����) Then
                                .Row = i: .Col = AI_����ҩ��
                                Call AddErrInfo("����" & IIf(blnDateIsNull, "���ڹ���ʱ��Ϊ�յ���ͬ�Ĺ���ҩ���¼��", Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��"), 0, vsTmp)
                            End If
                        End If
                    Next
                End If
            Next
        End With
        If blnJudge Then
            If gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼).Value = 0 Then
                Call AddErrInfo("�ò��˲����ڹ�����¼����û�й�ѡ�޹�����¼���Ƿ������", 1, gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼))
            End If
        End If
        blnJudge = False

        If gclsPros.PatiType <> PF_���� Then
            Set vsTmp = gclsPros.CurrentForm.vsOPS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, PI_��������)) <> "" Then
                        If .TextMatrix(i, PI_��������) = "" And gclsPros.FuncType = f������ҳ Then
                            .Row = i: .Col = PI_��������
                            Call AddErrInfo("�������������롣", 0, vsTmp)
                        End If
                        If Not IsDate(.TextMatrix(i, PI_��������)) Then
                            .Row = i: .Col = PI_��������
                            Call AddErrInfo("�����������벻��ȷ��", 0, vsTmp)
                        ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd") > Format(gclsPros.OutTime, "yyyy-MM-dd") Or _
                            Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                            .Row = i: .Col = PI_��������    '��������û�о�ȷ��ʱ��
                            Call AddErrInfo("�������ڲ������Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                        End If

                        If gclsPros.UseOPSEndTime Then
                            If Not IsDate(.TextMatrix(i, PI_��������)) Then
                                .Row = i: .Col = PI_��������
                                Call AddErrInfo("��������ʱ�����벻��ȷ��", 0, vsTmp)
                            ElseIf Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd HH:mm") < Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_��������
                                Call AddErrInfo("��������ʱ��������������ʼʱ�䡣", 0, vsTmp)
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                Format(.TextMatrix(i, PI_��������), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_��������
                                Call AddErrInfo("��������ʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                            End If
                        End If
                        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                            If Not IsDate(.TextMatrix(i, PI_����ʼʱ��)) Then
                                If .TextMatrix(i, PI_����ʼʱ��) <> "" Then
                                    .Row = i: .Col = PI_����ʼʱ��
                                    Call AddErrInfo("����ʼʱ�����벻��ȷ��", 0, vsTmp)
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_����ʼʱ��), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                Format(.TextMatrix(i, PI_����ʼʱ��), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_��������
                                Call AddErrInfo("����ʼʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                            End If

                            If Not IsDate(.TextMatrix(i, PI_������ҩʱ��)) And .TextMatrix(i, PI_������ҩʱ��) <> "" Then
                                .Row = i: .Col = PI_������ҩʱ��
                                Call AddErrInfo("������ҩʱ�����벻��ȷ��", 0, vsTmp)
                            End If
                            strSql = "Select ׼������ From ���������¼ Where Rownum = 1"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ֶγ���")
                            If Len(Trim(.TextMatrix(i, PI_׼������))) > 3 Then
                                Call AddErrInfo("׼���������벻��ȷ,�����ֵ������������Ƴ��ȡ�", 0, vsTmp)
                            End If
                        End If
                        If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                            If Len(Trim(.TextMatrix(i, PI_����ҩ����))) > 5 Then
                                Call AddErrInfo("������ҩ�������벻��ȷ,�����ֵ������������Ƴ��ȡ�", 0, vsTmp)
                            End If
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, PI_��������)) > 300 Then
                            .Row = i: .Col = PI_��������
                            Call AddErrInfo("������������̫����ֻ����300���ַ���150�����֡�", 0, vsTmp)
                        End If

                        If .ColHidden(PI_������ʿ) Then
                            If .TextMatrix(i, PI_����ҽʦ) = "" Then
                                .Row = i: .Col = PI_����ҽʦ
                                Call AddErrInfo("����������ҽʦ��", 0, vsTmp)
                            End If
                        Else
                            If .TextMatrix(i, PI_����ҽʦ) = "" And .TextMatrix(i, PI_������ʿ) = "" Then
                                .Row = i: .Col = PI_����ҽʦ
                                Call AddErrInfo("����������ҽʦ��������ʿ��", 0, vsTmp)
                            End If
                        End If
                        For j = i + 1 To .Rows - 1
                            If Trim(.TextMatrix(j, PI_��������)) <> "" Then
                            
                                '�����������������ڽ������в���ʾʱ���������ھ͵�����������
                                If gclsPros.MedPageSandard = ST_��������׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                                    If Not gclsPros.UseOPSEndTime Then
                                        .TextMatrix(i, PI_��������) = .TextMatrix(i, PI_��������)
                                        .TextMatrix(j, PI_��������) = .TextMatrix(j, PI_��������)
                                    End If
                                ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                                    .TextMatrix(i, PI_��������) = .TextMatrix(i, PI_��������)
                                    .TextMatrix(j, PI_��������) = .TextMatrix(j, PI_��������)
                                End If
                            
                                If gclsPros.MedPageSandard = ST_��������׼ Then
                                    If .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������ID) & "|" & .TextMatrix(i, PI_������ĿID) & "|" & .TextMatrix(i, PI_�пڲ�λ) = .TextMatrix(j, PI_��������) & "|" & .TextMatrix(j, PI_��������) & "|" & .TextMatrix(j, PI_��������) & "|" & .Cell(flexcpData, j, PI_��������) & "|" & .TextMatrix(j, PI_��������ID) & "|" & .TextMatrix(j, PI_������ĿID) & "|" & .TextMatrix(j, PI_�пڲ�λ) Then
                                        .Row = j: .Col = PI_��������
                                        Call AddErrInfo("���ִ��������������ڡ������������пڲ�λ����ͬ��������¼��", 0, vsTmp)
                                    End If
                                Else
                                    If .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������ID) & "|" & .TextMatrix(i, PI_������ĿID) = .TextMatrix(j, PI_��������) & "|" & .TextMatrix(j, PI_��������) & "|" & .TextMatrix(j, PI_��������) & "|" & .Cell(flexcpData, j, PI_��������) & "|" & .TextMatrix(j, PI_��������ID) & "|" & .TextMatrix(j, PI_������ĿID) Then
                                        .Row = j: .Col = PI_��������
                                        Call AddErrInfo("���ִ�����������������������������ͬ��������¼��", 0, vsTmp)
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            End With
        End If

        If gclsPros.FuncType = f������ҳ Then
            grsBabyInfo.Filter = "���䷽ʽ<>'��������'"
            If grsBabyInfo.RecordCount > 0 Then
                If Not gclsPros.Have���� Then
                    Call AddErrInfo("̥���ķ��䷽ʽ���ڲ����������䣬��δ������ص�����������Ƿ������", 1, gclsPros.CurrentForm.vsOPS)
                End If
            Else
                grsBabyInfo.Filter = "���䷽ʽ='��������'"
                If grsBabyInfo.RecordCount > 0 Then
                    If gclsPros.Have���� Then
                        Call AddErrInfo("̥���ķ��䷽ʽ���������䣬����������ص�����������Ƿ������", 1, gclsPros.CurrentForm.vsOPS)
                    End If
                End If
            End If

            Set vsTmp = gclsPros.CurrentForm.vsFees
            With vsTmp
                i = 3
                Do
                    On Error Resume Next
                    sgnTmp = Fix(Val(.TextMatrix(i \ 3, (i Mod 3) * 2 + 1)))
                    If Err.Number <> 0 Or Len(sgnTmp) > 11 Then
                        .Row = i \ 3: .Col = (i Mod 3) * 2 + 1
                        Call AddErrInfo("���ý����ֵ̫��", 0, vsTmp)
                        Err.Clear: On Error GoTo 0
                    End If
                    If Val(.TextMatrix(i \ 3, (i Mod 3) * 2 + 1)) <> 0 Then
                        If .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*����*" And Not .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*������*" And Not gclsPros.Have���� Then
                            gclsPros.CurrentForm.vsOPS.Row = gclsPros.CurrentForm.vsOPS.FixedRows: gclsPros.CurrentForm.vsOPS.Col = PI_��������
                            Call AddErrInfo("�ò���סԺ�����к��������ѣ���û��¼��������Ϣ���Ƿ������", 1, gclsPros.CurrentForm.vsOPS)
                        ElseIf .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*��Ѫ*" Or .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*Ѫ��*" Then
                            If gclsPros.CurrentForm.cboBaseInfo(BCC_Ѫ��).Text = "" And gclsPros.CurrentForm.cboBaseInfo(BCC_RH).Text = "" And _
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_���ϸ��).Text = "" And gclsPros.CurrentForm.txtSpecificInfo(SLC_��ѪС��).Text = "" And _
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_��Ѫ��).Text = "" And gclsPros.CurrentForm.txtSpecificInfo(SLC_��ȫѪ).Text = "" And _
                                gclsPros.CurrentForm.txtInfo(GC_������).Text = "" Then
                                Call AddErrInfo("�ò��˴�����Ѫ�ѣ���ѡ��Ѫ�͡�Rh���������ϸ����" & vbCrLf & "��ѪС�塢��Ѫ������ȫѪ������������Ӧ���", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_Ѫ��), gclsPros.CurrentForm.cboBaseInfo(BCC_RH))
                            End If
                        End If
                    End If
                    j = i + 1
                    If j <= .Rows * 3 - 1 Then
                        Do
                            If .TextMatrix(j \ 3, (j Mod 3) * 2) <> "" Then
                                If GetTextByDot(.TextMatrix(i \ 3, (i Mod 3) * 2)) = GetTextByDot(.TextMatrix(j \ 3, (j Mod 3) * 2)) Then
                                    If Not gclsPros.SameName Then
                                        .Row = j \ 3: .Col = (j Mod 3) * 2
                                        Call AddErrInfo("������ñ��С�" & GetTextByDot(.TextMatrix(i \ 3, (i Mod 3) * 2)) & "�����������˶�Ρ�", 0, vsTmp)
                                    Else '�ϲ���������
                                        .TextMatrix(i \ 3, (i Mod 3) * 2) = Format(Val(.TextMatrix(i \ 3, (i Mod 3) * 2)) + Val(.TextMatrix(j \ 3, (j Mod 3) * 2)), gclsPros.FreeFormat)
                                        Call AddOrDelFreeCols(vsTmp, .TextMatrix(j \ 3, (j Mod 3) * 2), .TextMatrix(j \ 3, (j Mod 3) * 2 + 1), False)
                                    End If
                                End If
                            End If
                            If j < .Rows * 3 - 1 Then
                                j = j + 1: blnDoEx = True
                            Else
                                blnDoEx = False
                            End If
                        Loop While blnDoEx
                    End If
                    If i < .Rows * 3 - 1 Then
                        i = i + 1: blnDo = True
                    Else
                        blnDo = False
                    End If
                Loop While blnDo
            End With
        End If
        If gclsPros.PatiType <> PF_���� Then
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_����ҩ����) <> "" Then
                        If Trim(.TextMatrix(i - 1, KI_����ҩ����)) = "" Then
                            .Row = i - 1: .Col = KI_����ҩ����
                            Call AddErrInfo("���������뿹��ҩ�����ݡ�", 0, vsTmp)
                        End If
                        If (Len(.TextMatrix(i, KI_ʹ������)) > 18 Or Val(.TextMatrix(i, KI_ʹ������)) = 0) And Trim(.TextMatrix(i, KI_ʹ������)) <> "" Then
                            .Row = i: .Col = KI_ʹ������
                            Call AddErrInfo("����дʮ��λ�����ڵ�����������", 0, vsTmp)
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, KI_��ҩĿ��)) > 200 And Trim(.TextMatrix(i, KI_��ҩĿ��)) <> "" Then
                            .Row = i: .Col = KI_��ҩĿ��
                            Call AddErrInfo("����д100���������ڵ���ҩĿ�ġ�", 0, vsTmp)
                        End If

                        For j = .FixedRows To i - 1
                            If Trim(.TextMatrix(j, KI_����ҩ����)) = Trim(.TextMatrix(i, KI_����ҩ����)) And Trim(.TextMatrix(j, KI_��ҩĿ��)) = Trim(.TextMatrix(i, KI_��ҩĿ��)) And Trim(.TextMatrix(j, KI_ʹ�ý׶�)) = Trim(.TextMatrix(i, KI_ʹ�ý׶�)) Then
                                .Row = j: .Col = KI_����ҩ����
                                Call AddErrInfo("���ִ���������ͬ�Ŀ���ҩ����Ϣ��", 0, vsTmp)
                            End If
                        Next
                    End If
                Next
            End With
            If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                Set vsTmp = gclsPros.CurrentForm.vsTSJC
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, 1)) <> "" Then
                            If i > .FixedRows Then
                                If Trim(.TextMatrix(i - 1, 1)) = "" Then
                                    .Row = i - 1: .Col = 1
                                    Call AddErrInfo("�������������������ݡ�", 0, vsTmp)
                                End If
                            End If

                            For j = .FixedRows To i - 1
                                If Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                                    .Row = j: .Col = 1
                                    Call AddErrInfo("���ִ���������ͬ����������Ϣ��", 0, vsTmp)
                                End If
                            Next
                        End If
                    Next
                End With
            End If
            If gclsPros.MedPageSandard = ST_��������׼ Or gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, UI_�໤������)) <> "" Then
                            If zlCommFun.ActualLen(.TextMatrix(i, UI_�໤������)) > 100 Then
                                .Row = i: .Col = UI_�໤������
                                Call AddErrInfo("��֢�໤��������������̫����ֻ����100���ַ���50�����֡�", 0, vsTmp)
                            End If
                            If Trim(.TextMatrix(i, UI_����ʱ��)) <> "____-__-__ __:__" Then
                                If Not IsDate(.TextMatrix(i, UI_����ʱ��)) Then
                                     .Row = i: .Col = UI_����ʱ��
                                    If gclsPros.MedPageSandard = ST_��������׼ Then
                                         Call AddErrInfo("����ʱ�����벻��ȷ��", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                                         Call AddErrInfo("��סʱ�����벻��ȷ��", 0, vsTmp)
                                    End If
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, UI_����ʱ��), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                     Format(.TextMatrix(i, UI_����ʱ��), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                    .Row = i: .Col = UI_����ʱ��
                                    If gclsPros.MedPageSandard = ST_��������׼ Then
                                        Call AddErrInfo("����ʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                                        Call AddErrInfo("��סʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                                    End If
                            End If
                            If Trim(.TextMatrix(i, UI_�˳�ʱ��)) <> "____-__-__ __:__" Then
                                If Not IsDate(.TextMatrix(i, UI_�˳�ʱ��)) Then
                                    .Row = i: .Col = UI_�˳�ʱ��
                                    If gclsPros.MedPageSandard = ST_��������׼ Then
                                        Call AddErrInfo("�˳�ʱ�����벻��ȷ��", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                                        Call AddErrInfo("ת��ʱ�����벻��ȷ��", 0, vsTmp)
                                    End If
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, UI_�˳�ʱ��), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                     Format(.TextMatrix(i, UI_�˳�ʱ��), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                    .Row = i: .Col = UI_�˳�ʱ��
                                    If gclsPros.MedPageSandard = ST_��������׼ Then
                                        Call AddErrInfo("�˳�ʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                                        Call AddErrInfo("ת��ʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                                    End If
                            End If
                            If Trim(.TextMatrix(i, UI_�˳�ʱ��)) <> "" And Trim(.TextMatrix(i, UI_����ʱ��)) <> "" And CDate(Trim(.TextMatrix(i, UI_�˳�ʱ��))) < CDate(Trim(.TextMatrix(i, UI_����ʱ��))) Then
                                .Row = i: .Col = UI_����ʱ��
                                If gclsPros.MedPageSandard = ST_��������׼ Then
                                    Call AddErrInfo("����ICU��ʱ�����С���˳�ICU��ʱ�䡣", 0, vsTmp)
                                ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                                    Call AddErrInfo("��סICU��ʱ�����С��ת��ICU��ʱ�䡣", 0, vsTmp)
                                End If
                            End If
                        End If
                    Next
                End With
            End If

            '��֢�໤��е��ҽԺ��Ⱦ���걾�����飬���Ĵ�����
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, TI_ICU����) <> "" And .TextMatrix(i, TI_��е������) <> "" Then
                            j = Val(.Cell(flexcpData, i, TI_ICU����))
                            str��סʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(j, UI_����ʱ��))
                            strת��ʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(j, UI_�˳�ʱ��))
                            If zlCommFun.ActualLen(.TextMatrix(i, TI_ICU����)) > 50 Then
                                .Row = i: .Col = TI_ICU����
                                Call AddErrInfo("ICU���͵����ݵ����ֻ����50���ַ�/25�����֡�", 0, vsTmp)
                            End If
                            If Not IsDate(.TextMatrix(i, TI_��ʼʱ��)) Then
                                .Row = i: .Col = TI_��ʼʱ��
                                Call AddErrInfo("��������ȷ�Ŀ�ʼʹ��ʱ�䡣", 0, vsTmp)
                            Else
                                If IsDate(str��סʱ��) Then
                                    If CDate(.TextMatrix(i, TI_��ʼʱ��)) < CDate(str��סʱ��) Then
                                        .Row = i: .Col = TI_��ʼʱ��
                                        Call AddErrInfo("��ʼʹ��ʱ��С������֢�໤��������סʱ��,����", 0, vsTmp)
                                    End If
                                End If
                            End If
                            If Not IsDate(.TextMatrix(i, TI_����ʱ��)) Then
                                .Row = i: .Col = TI_����ʱ��
                                Call AddErrInfo("��������ȷ�Ŀ�ʼʹ��ʱ�䡣", 0, vsTmp)
                            Else
                                If IsDate(strת��ʱ��) Then
                                    If CDate(.TextMatrix(i, TI_����ʱ��)) > CDate(strת��ʱ��) Then
                                        .Row = i: .Col = TI_����ʱ��
                                        Call AddErrInfo("����ʹ��ʱ���������֢�໤������ת��ʱ��,����", 0, vsTmp)
                                    End If
                                End If
                            End If
                            If IsDate(.TextMatrix(i, TI_��ʼʱ��)) And IsDate(.TextMatrix(i, TI_����ʱ��)) Then
                                If CDate(.TextMatrix(i, TI_��ʼʱ��)) > CDate(.TextMatrix(i, TI_����ʱ��)) Then
                                    .Row = i: .Col = TI_��ʼʱ��
                                    Call AddErrInfo("��ʼʹ��ʱ������˽���ʹ��ʱ�䣬���顣", 0, vsTmp)
                                End If
                            End If
                        End If
                    Next
                End With

                Set vsTmp = gclsPros.CurrentForm.vsInfect
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If IsDate(.TextMatrix(i, FI_ȷ������)) Then
                            If gclsPros.OutTime <> "" And Format(.TextMatrix(i, FI_ȷ������), "yyyy-MM-dd") > Format(gclsPros.OutTime, "yyyy-MM-dd") Or _
                                Format(.TextMatrix(i, FI_ȷ������), "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                                .Row = i: .Col = FI_ȷ������
                                Call AddErrInfo("ȷ��ʱ�䲻�����Ժ���ڷ�Χ�ڡ�", 0, vsTmp)
                            End If
                            If .TextMatrix(i, FI_��Ⱦ��λ) = "" Then
                                .Row = i: .Col = FI_��Ⱦ��λ
                                Call AddErrInfo("�������Ⱦ��λ��", 0, vsTmp)
                            End If
                            If .TextMatrix(i, FI_ҽԺ��Ⱦ����) = "" Then
                                .Row = i: .Col = FI_ҽԺ��Ⱦ����
                                Call AddErrInfo("������ҽԺ��Ⱦ���ơ�", 0, vsTmp)
                            End If
                        ElseIf Trim(.TextMatrix(i, FI_ȷ������)) <> "" And Not IsDate(.TextMatrix(i, FI_ȷ������)) Then
                            .Row = i: .Col = FI_ȷ������
                            Call AddErrInfo("��������ȷ��ȷ�����ڡ�", 0, vsTmp)
                        End If
                    Next
                End With

                Set vsTmp = gclsPros.CurrentForm.vsSample
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, MI_�걾) <> "" Then
                            If .TextMatrix(i, MI_��ԭѧ���뼰����) = "" Then
                                .Row = i: .Col = MI_��ԭѧ���뼰����
                                Call AddErrInfo("�����벡ԭѧ���뼰���ơ�", 0, vsTmp)
                            End If
                            If Not IsDate(.TextMatrix(i, MI_�ͼ�����)) Then
                                .Row = i: .Col = MI_�ͼ�����
                                Call AddErrInfo("��������ȷ���ͼ����ڡ�", 0, vsTmp)
                            End If
                        End If
                    Next
                End With

            End If

            If gclsPros.ReadPages And gclsPros.MedPageSandard = ST_��������׼ Then
                Set vsTmp = gclsPros.CurrentForm.vsSpirit
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, SI_ҩ������)) <> "" Then
                            '�����е����������飬��ΪForm_KeyDown�¼��Ѿ�����
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_ҩ������))) > 200 Then
                                .Row = i: .Col = SI_ҩ������
                                Call AddErrInfo("ҩ��������������̫����ֻ����200���ַ���100�����֡�", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_�Ƴ�))) > 50 Then
                                .Row = i: .Col = SI_�Ƴ�
                                 Call AddErrInfo("�Ƴ���������̫����ֻ����50���ַ���25�����֡�", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_��Ч))) > 50 Then
                                .Row = i: .Col = SI_��Ч
                                Call AddErrInfo("��Ч��������̫����ֻ����50���ַ���25�����֡�", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_���ⷴӦ))) > 100 Then
                                .Row = i: .Col = SI_���ⷴӦ
                                Call AddErrInfo("���ⷴӦ��������̫����ֻ����100���ַ���50�����֡�", 0, vsTmp)
                            End If

                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_�������))) > 50 Then
                                .Row = i: .Col = SI_�������
                                Call AddErrInfo("���������������̫����ֻ����50���ַ���25�����֡�", 0, vsTmp)
                            End If
                        End If
                    Next
                End With
            End If
        End If
    End If
    
    
   '����Ƿ�����Ҳ��������ҳ
    Call CreatePlugInOK(gclsPros.Module)
    If Not gobjPlugIn Is Nothing Then
        Err.Clear: On Error Resume Next
        If gobjPlugIn.gblnmec = True Then
            '���ò������ӿ�
            If Err.Number = 0 Then
                Set gColCtl = CtlAdd
                strMsg = ""
                If gobjPlugIn.CheckMecInfo(gclsPros.SysNo, gclsPros.Module, gclsPros.����ID, gclsPros.��ҳID, gColCtl, strMsg) = False Then
                    If strMsg <> "" And Err.Number = 0 Then
                        Call ErrDw(strMsg)
                    End If
                End If
            End If
            Call zlPlugInErrH(Err, "CheckMecInfo")
            Err.Clear: On Error GoTo 0
        End If
    End If
    
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Err.Clear: On Error Resume Next
            Call gfrmMecCol(i).CheckPlugMec(gclsPros.SysNo, gclsPros.Module, gclsPros.����ID, gclsPros.��ҳID, colErr)
            Call zlPlugInErrH(Err, "CheckPlugMec")
        Next
        If colErr.Count > 0 And Err.Number = 0 Then
            Call ErrMec(colErr)
            Set colErrTmp = colErr
        End If
        Err.Clear: On Error GoTo 0
    End If
     
'    ���ش���;��浽������
    If gColErr.Count > 0 Or gColWarn.Count > 0 Then
        Call LoadVsErrData
        If Not blnCheck Then
            If gColErr.Count = 0 And gColWarn.Count > 0 Then
                If MsgBox("����" & CStr(gColWarn.Count) & "�����棬�Ƿ����ȫ�����棬����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    Call ClearErrCol
                    If blnBaseInfo Then
                        With gclsPros.CurrentForm
                            If Not gobjPatient.SavePatiBaseInfo(gclsPros.����ID, gclsPros.��ҳID, .txtInfo(GC_����).Text, strSex, strAge, strBirthday, IIf(gclsPros.PatiType = PF_����, "������ҳ", "סԺ��ҳ"), gclsPros.PatiType, strErrIfno) Then
                                If InStr(strErrIfno, "�Ա�") > 0 Then
                                    Set objTmp = .cboBaseInfo(BCC_�Ա�)
                                ElseIf InStr(strErrIfno, "��������") > 0 Then
                                    Set objTmp = .mskDateInfo(DC_��������)
                                ElseIf InStr(strErrIfno, "����") > 0 Then
                                    Set objTmp = .txtSpecificInfo(SLC_����)
                                End If
                                If ShowMessage(objTmp, "���֤�����ȡ��" & strBaseInfo & "�뵱ǰ�����" & strBaseInfo & "��������Զ����½����ϵ�" & strBaseInfo & "ʧ�ܣ�ʧ��ԭ��" & strErrIfno & ",�Ƿ������", True) = vbNo Then Exit Function
                            End If
                            Call SetCtrlValues("�Ա�", strSex)
                            Call SetCtrlValues("����", strAge)
                            Call SetCtrlValues("��������", strBirthday)
                        End With
                    End If
                    CheckMedPageData = True
                End If
            ElseIf gColErr.Count > 0 Then
                CheckMedPageData = False
                Exit Function
            End If
        End If
    End If
    
    If Not CheckMedPageChange Then
        gclsPros.InfosChange = False
        gclsPros.IsCheckData = False
        Exit Function
    Else
        gclsPros.IsCheckData = False
    End If
    
    CheckMedPageData = True
    Exit Function
errH:
    Debug.Print "CheckMedPageData:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckMedPageChange() As Boolean
'���ܣ������ҳ��Ϣ�Ƿ����仯
'�����������
    On Error GoTo errH
'    gclsPros.InfosChange = False
    With gclsPros.CurrentForm
        Call gclsPros.RollBackCacheRecInfo '�ع���Ϣ����
        '������ҳ��������ҳ�ӱ���Ϣ����
        Call CacheCtrlValues
        '��ϻ���
        '1�����没ԭѧ���
        If gclsPros.PatiType <> PF_���� Then Call UpdateCacheRecInfo(1, "��ԭѧ���", .txtInfo(GC_��ԭѧ���).Text & .cmdInfo(GC_��ԭѧ���).Tag)
        '2��������ҽ���
        Call CacheLoadVsDiagData(.vsDiagXY, , , True)
        '3��������ҽ���
        If gclsPros.Have��ҽ Then
            Call CacheLoadVsDiagData(.vsDiagZY, , , True)
        End If
        '������Ϣ����
        Call CacheLoadVsAllerData(.vsAller, , True)
        If gclsPros.PatiType <> PF_���� Then
            '��������
            Call CacheLoadVsOPSData(.vsOPS, , True)
            '��Ϸ����������
            Call CacheLoadDiagMatchData(, True)
            '���˷�����Ϣ����
            If gclsPros.FuncType = f������ҳ Then
                Call CacheLoadVsFreesData(.vsFees, , True)
            End If
            '����ҩʹ���������
            Call CacheLoadVsKSSData(.vsKSS, , True)
            '��֢�໤ʹ���������
            If gclsPros.MedPageSandard <> ST_����ʡ��׼ Then
                If gclsPros.MedPageSandard <> ST_����ʡ��׼ Then
                    Call CacheLoadVsFlxAddICUData(.vsFlxAddICU, , True)
                Else
                    Call CacheLoadVsFlxAddICUData(, , True)
                End If
            End If
            '���ơ����ơ�����ҩƷ����
            If gclsPros.ReadPages Then
                Call CacheLoadVsChemothData(.vsChemoth, , True)
                Call CacheLoadVsRadiothData(.vsRadioth, , True)
                If gclsPros.MedPageSandard = ST_��������׼ Then Call CacheLoadVsSpiritData(.vsSpirit, , True)
            End If
            '��֢�໤��е��ҽԺ��Ⱦ���걾�������
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Call CacheLoadVsICUInstrumentsData(.vsICUInstruments, , True)
                Call CacheLoadvsInfectData(.vsInfect, , True)
                Call CacheLoadvsSampleData(.vsSample, , True)
            End If
        End If
        '������ؼ�ֵ�仯
        Call UpdateCacheRecInfo(2)
        '������鲡���Ƿ��Ŀ����ҳ��������״̬
        If gclsPros.PatiType <> PF_���� And gclsPros.FuncType <> f������ҳ Then
            If Not CheckMecRed(gclsPros.����ID, gclsPros.��ҳID, .Caption, "�޸���ҳ") Then Exit Function
        End If
    End With
    CheckMedPageChange = True
    Exit Function
errH:
    Debug.Print "CheckMedPageChange:" & Err.Source & "===" & Err.Description
    Call gclsPros.RollBackCacheRecInfo '�ع���Ϣ����
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SaveMedPageData() As Boolean
'���ܣ����没����ҳ����
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim datCur As Date, strMsg As String
    Dim blnChange As Boolean
    
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            blnChange = blnChange Or gfrmMecCol(i).gblnchange
        Next
        blnChange = blnChange Or gclsPros.InfosChange
    Else
        blnChange = gclsPros.InfosChange
    End If
    If blnChange Then
        datCur = zlDatabase.Currentdate
        arrSQL = Array()
        '�ֽ��в�����ҳ�Լ�������Ϣ�ĸ��£���ΪZL_������ҳ_��ҳ�����ԭʼ������ҳ�ӱ��еġ����Σ�סԺ�����Σ����ҽʦ����Ϣ���ж�ȡ
        '���ȱ��没����ҳ�ӱ���ᵼ�£���ȡ�����º����Ϣ������ZL_������ҳ_��ҳ������ʧ��
        gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
        If Not gclsPros.Is��ʿվ And gclsPros.MainInfoRec.RecordCount <> 0 Then
            Call PopPatiMainSQL(arrSQL)
        End If
        '�ӱ���Ϣ����
        Call PopPatiAuxiSQL(arrSQL, gclsPros.Is��ʿվ)
        'ҽ���뻤ʿ������ҳ,ҽ��վ������Ĵӱ���Ŀ
        If Not gclsPros.SeparateEdit And gclsPros.PatiType = PF_סԺ Then
            Call PopPatiAuxiSQL(arrSQL, True)
        End If
        If Not gclsPros.Is��ʿվ Then
            gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
            If gclsPros.MainInfoRec.RecordCount <> 0 Then
                If gclsPros.PatiType = PF_סԺ Then
                    '�ṹ����ַ����
                    If gclsPros.IsStructAdress Then
                        Call PopStructAdressSQL(arrSQL)
                    End If
                    '��Ϸ����������
                    Call PopDiagMatchSQL(arrSQL)
                    '���鱣��
                    Call PopOPSSQL(arrSQL)
                    '����ҩ����
                    Call PopKSSSQL(arrSQL)
                    If gclsPros.MedPageSandard <> ST_����ʡ��׼ Then
                        '��֢�໤�������
                        Call PopICUSQL(arrSQL)
                        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then Call PopOtherSQL(arrSQL)
                    End If
                    '���˷�����Ϣ����
                    If gclsPros.FuncType = f������ҳ Then
                        Call PopFeeSQL(arrSQL)
                    End If
                    '���ƣ����ƣ�����ҩ����
                    If gclsPros.ReadPages Then
                        Call PopShareInfoSQL(arrSQL)
                    End If
                End If
                '��������
                Call PopAllerSQL(arrSQL)
                '��ϱ���
                Call PopPatiDiagSQL(arrSQL, datCur)
            End If
            '������Ϣ����
            If gclsPros.FuncType = f������ҳ Then Call PopDelicerySQL(arrSQL)
        End If
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), gclsPros.CurrentForm.Caption)
        Next
        
        '��Ҹ�ҳ���ݱ���
        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
            For i = 1 To gfrmMecCol.Count
                If gfrmMecCol(i).gblnchange Then
                    Err.Clear: On Error Resume Next
                    If gfrmMecCol(i).savePlugMec(gclsPros.SysNo, gclsPros.Module, gclsPros.����ID, gclsPros.��ҳID) = False Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                    Call zlPlugInErrH(Err, "SavePlugMec")
                    Err.Clear: On Error GoTo 0
                End If
            Next
        End If


        If gclsPros.FuncType = fҽ����ҳ Then
            If gclsPros.PatiType = PF_���� Then
                '��������ͬ��
                If Not gobjCommunity Is Nothing And gclsPros.CommunityID <> 0 Then
                    If Not gobjCommunity.UpdateInfo(gclsPros.SysNo, p����ҽ��վ, gclsPros.CommunityID, gclsPros.CommunityNO, gclsPros.����ID, gclsPros.��ҳID) Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                End If
            Else
                '����ҽ��������Ϣ�޸Ľӿ�
                If gclsPros.InsureType <> 0 And Not gclsInsure Is Nothing Then
                    If Not gclsInsure.ModiPatiSwap(gclsPros.����ID, gclsPros.��ҳID, gclsPros.InsureType, "2") Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                End If
            End If
            
            If gobjPlugIn Is Nothing Then
                Call CreatePlugInOK(IIf(gclsPros.PatiType = PF_����, p����ҽ��վ, pסԺҽ��վ))
            End If
            If Not gobjPlugIn Is Nothing And gclsPros.PatiType = PF_סԺ Then
                Err.Clear: On Error Resume Next
                If gobjPlugIn.EMPI_ModifyPatiInfo(gclsPros.SysNo, pסԺҽ��վ, gclsPros.����ID, gclsPros.��ҳID, 0, strMsg) = 0 Then
                    If Err.Number = 0 Then
                        gcnOracle.RollbackTrans
                        Screen.MousePointer = 0
                        MsgBox "��ǰ������EMPIϵͳ�ӿڣ���EMPIϵͳ�ӿ�(EMPI_ModifyPatiInfo)δ���óɹ�:" & strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                If Err.Number <> 0 And Err.Number <> 438 Then
                    gcnOracle.RollbackTrans
                    Screen.MousePointer = 0
                    Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo")
                    Exit Function
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
        gcnOracle.CommitTrans: blnTrans = False
        '��Ϣ����
        If gclsPros.FuncType <> f������ҳ Then Call SendMsgDiag(datCur)
        '�����ӿڵ���
        If HaveRIS Then
            Call gobjRis.HISModPati(gclsPros.PatiType, gclsPros.����ID, gclsPros.��ҳID)
        End If
    End If
    '������Ϣ��¼�������£�����ֵ��ֵ��ԭֵ,����ʼ���ı�״̬
    'Ŀ�ģ���ҳ���ڴ�ӡԤ�����ܣ��������Ա༭���༭���ֿ��Ա���
    
    Call gclsPros.InitCacheRecInfo(True)
    
    If gclsPros.OpenMode <> EM_�������� And gclsPros.OpenMode <> EM_������ҳ Then
        On Error Resume Next
        Call LoadDiagAndAllerFData
    End If

    On Error GoTo errH
    Screen.MousePointer = 0
    gclsPros.InfosChange = False
    SaveMedPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SendMsgDiag(ByVal datCur As Date) As Boolean
'���ܣ����������Ϣ
    Dim i As Long
    Dim arrTmp As Variant
    Dim strFilter As String
    On Error GoTo errH
    If gclsMipModule Is Nothing Then SendMsgDiag = True: Exit Function
    gclsPros.MainInfoRec.Filter = "(��Ϣ��='��ҽ���' And �Ƿ�ı�=1) OR (��Ϣ��='��ҽ���' And �Ƿ�ı�=1) "
    For i = 1 To gclsPros.MainInfoRec.RecordCount
        gclsPros.SecdInfoRec.Filter = "�ı�״̬<>" & CS_δ�ı� & " And �ı�״̬<>" & CS_������ & "   And ���=" & gclsPros.MainInfoRec!���
        Do While Not gclsPros.SecdInfoRec.EOF
            arrTmp = Split(gclsPros.SecdInfoRec!��Ϣԭֵ & "", "|")
            If gclsPros.SecdInfoRec!�ı�״̬ <> CS_������ Then 'ɾ�������滻���ȴ���ɾ�������Ϣ
                Call ZLHIS_CIS_011(gclsMipModule, gclsPros.����ID, gclsPros.PatiName, gclsPros.PatiType, gclsPros.��ҳID, gclsPros.��Ժ����ID, gclsPros.SecdInfoRec!ID, arrTmp(DMP_��ϱ���), arrTmp(DMP_��������))
            End If
            arrTmp = Split(gclsPros.SecdInfoRec!��Ϣ��ֵ & "", "|")
            If gclsPros.SecdInfoRec!�ı�״̬ <> CS_ɾ���� Then  '���������滻�д����´������Ϣ
                Call ZLHIS_CIS_010(gclsMipModule, gclsPros.����ID, gclsPros.PatiName, gclsPros.PatiType, gclsPros.��ҳID, gclsPros.��Ժ����ID, Val(gclsPros.SecdInfoRec!Tag & ""), arrTmp(DMP_�������), arrTmp(DMP_�Ƿ�����), arrTmp(DMP_��ϴ���), arrTmp(DMP_��ϱ���), arrTmp(DMP_��������), arrTmp(DMP_��������), arrTmp(DMP_�������), arrTmp(DMP_֤�����), arrTmp(DMP_֤������), datCur, UserInfo.����)
            End If
            gclsPros.SecdInfoRec.MoveNext
        Loop
    Next
    SendMsgDiag = True
    '��ԭѧ��ϲ�������Ϣ
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PopPatiMainSQL(ByRef arrSQL As Variant)
'���ܣ�����ȡ�Ĳ�����ҳ������Ϣ��SQL����SQL������
'SaveMedPageData���Ӻ���
    Dim lngCallProc As Long
    Dim arrField As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    arrField = Array()
      '������ҳ��������Ϣ���Լ�������ҳ�ӱ�
    '1��������ҳ�Լ�������Ϣ�ı���
    '���жϣ���Ҫ������Щ�洢����
    '������ҳ�벡����Ϣ������Ŀ���
    gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
    If gclsPros.PatiType = PF_סԺ Then
        arrField = Array("����", "����", "����", "ְҵ", "����״��", "ҽ�Ƹ��ʽ", "��ͥ��ַ", "��ͥ�绰", "��ͥ��ַ�ʱ�", "��λ��ַ", _
                        "��λ�绰", "��λ�ʱ�", "��ϵ������", "��ϵ�˹�ϵ", "��ϵ�˵绰", "��ϵ�˵�ַ", "���ڵ�ַ", "���ڵ�ַ�ʱ�")
        If gclsPros.FuncType = f������ҳ Then
             ReDim Preserve arrField(UBound(arrField) + 3)
             arrField(UBound(arrField) - 2) = "סԺ��"
             arrField(UBound(arrField) - 1) = "��Ժ����"
             arrField(UBound(arrField)) = "��Ժ����"
        End If

        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.MoveFirst
            For j = 1 To gclsPros.MainInfoRec.RecordCount
                If arrField(i) = gclsPros.MainInfoRec!��Ϣ�� Then
                    lngCallProc = 1: Exit For
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
            If lngCallProc = 1 Then Exit For
        Next
    End If
    If lngCallProc <> 1 Then
        If gclsPros.PatiType = PF_סԺ Then
            '������Ϣ������Ŀ���
            arrField = Array("����", "�Ա�", "����", "����", "��������", "�����ص�", "���֤��", "����֤��")
            If gclsPros.FuncType <> f������ҳ Then
                ReDim Preserve arrField(UBound(arrField) + 1)
                arrField(UBound(arrField)) = "סԺ��"
            End If
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                ReDim Preserve arrField(UBound(arrField) + 2)
                arrField(UBound(arrField)) = "Qq"
                arrField(UBound(arrField) - 1) = "Email"
            End If
        Else
             arrField = Array("�����", "����", "�Ա�", "����", "����", "����", "����", "����", "ְҵ", "��������", "�����ص�", "���֤��", _
                     "����֤��", "����״��", "ҽ�Ƹ��ʽ", "��ͥ��ַ", "��ͥ�绰", "��ͥ��ַ�ʱ�", "���ڵ�ַ", "���ڵ�ַ�ʱ�", "��ͬ��λID", "��λ��ַ", _
                     "��λ�绰", "��λ�ʱ�", "�໤��", "����", "ժҪ", "��Ⱦ���ϴ�", "����ʱ��", "������ַ")
        End If
        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.MoveFirst
            For j = 1 To gclsPros.MainInfoRec.RecordCount
                If arrField(i) = gclsPros.MainInfoRec!��Ϣ�� Then
                    lngCallProc = 2: Exit For
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
            If lngCallProc = 2 Then Exit For
        Next

        If gclsPros.PatiType = PF_סԺ Then
            '������ҳ������Ŀ���
            arrField = Array("��ҳid", "���", "����", "Ѫ��", "��Ժ����", "��Ժ��ʽ", "��Ժ��ʽ", "����Ժ", "�Ƿ�ȷ��", "ȷ������", "ʬ���־", "�����־", "��������", _
                    "�·�����", "��ҽ�������", "���ȴ���", "�ɹ�����", "����ҽʦ", "סԺҽʦ", "���λ�ʿ")
            If gclsPros.FuncType = f������ҳ Then
                arrTmp = Split("������,������,��Ժ����ID,��Ժ����ID,סԺ����,���ú�,��ĿԱ����,��Ŀ����", ",")
            Else
                arrTmp = Split("����ҽʦ,����ҽʦ,����Ա���,����Ա����", ",")
            End If
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = arrTmp(i)
            Next
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.MoveFirst
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!��Ϣ�� Then
                        lngCallProc = IIf(lngCallProc = 0, 3, 1): Exit For
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
                If lngCallProc = 1 Or lngCallProc = 3 Then Exit For
            Next
        End If
    End If
    If gclsPros.FuncType = f������ҳ Then
        '������������ģʽ��һ����ͬʱ���²�����ҳ�Լ�������Ϣ����������ģʽһ������²�����Ϣ
        lngCallProc = DecodeEx(gclsPros.OpenMode = EM_��������, 1, gclsPros.OpenMode = EM_������ҳ And lngCallProc = 2, 1, lngCallProc)
    End If
    If lngCallProc <> 0 Then
        'ZL_������Ϣ_��ҳ�������
        If lngCallProc <> 3 Then
            arrField = Array("����ID", IIf(gclsPros.PatiType = PF_����, "�����", "סԺ��"), "����", "�Ա�", "����", "����", "����", "����", "����", "ְҵ", "��������", "�����ص�", "���֤��", _
                        "����֤��", "����״��", "ҽ�Ƹ��ʽ", "��ͥ��ַ", "��ͥ�绰", "��ͥ��ַ�ʱ�", "���ڵ�ַ", "���ڵ�ַ�ʱ�", "��ͬ��λID", "��λ��ַ", "��λ�绰", "��λ�ʱ�")
            If gclsPros.FuncType = f������ҳ Then
                arrTmp = Split("��ϵ������,��ϵ�˹�ϵ, ��ϵ�˵绰, ��ϵ�˵�ַ," & IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, "Email,QQ,", ",,") & "��Ժ����,��Ժ����,סԺ����,��ҳID", ",")
            ElseIf gclsPros.PatiType = PF_סԺ Then
                arrTmp = Split("��ϵ������,��ϵ�˹�ϵ, ��ϵ�˵绰, ��ϵ�˵�ַ," & IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, "Email,QQ", ",,"), ",")
            Else
                arrTmp = Split(", , , ,,,�໤��,NO,����,ժҪ,��Ⱦ���ϴ�,����ʱ��,������ַ", ",")
            End If
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = arrTmp(i)
            Next
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = Get��ҳ����SQL(0, arrField) '��ȡZL_������Ϣ_��ҳ����ĵ���SQL
        End If

        'ZL_������ҳ_��ҳ�������
        If lngCallProc <> 2 Then
            If gclsPros.FuncType = f������ҳ Then
                arrField = Array("����ID", "��ҳID", "סԺ��", "������", "������", "����", "����", "����", "ְҵ", "���", "����", "Ѫ��", "����״��", "ҽ�Ƹ��ʽ", _
                    "��ͥ��ַ", "��ͥ�绰", "��ͥ��ַ�ʱ�", "���ڵ�ַ", "���ڵ�ַ�ʱ�", "��λ��ַ", "��λ�绰", "��λ�ʱ�", "��ϵ������", "��ϵ�˹�ϵ", _
                    "��ϵ�˵绰", "��ϵ�˵�ַ", "��Ժ����", "��Ժ��ʽ", "��Ժ����ID", "��Ժ����", "��Ժ��ʽ", "��Ժ����ID", "��Ժ����", "����Ժ", _
                    "�Ƿ�ȷ��", "ȷ������", "ʬ���־", "�����־", "��������", "�·�����", "��ҽ�������", "���ȴ���", "�ɹ�����", "סԺ����", _
                    "���ú�", "����ҽʦ", "סԺҽʦ", "���λ�ʿ", "��ĿԱ����", "��Ŀ����")
            Else
                arrField = Array("����ID", "��ҳID", "����", "����", "����", "ְҵ", "���", "����", "Ѫ��", "����״��", "ҽ�Ƹ��ʽ", _
                    "��ͥ��ַ", "��ͥ�绰", "��ͥ��ַ�ʱ�", "���ڵ�ַ", "���ڵ�ַ�ʱ�", "��λ��ַ", "��λ�绰", "��λ�ʱ�", "��ϵ������", "��ϵ�˹�ϵ", _
                    "��ϵ�˵绰", "��ϵ�˵�ַ", "��Ժ����", "��Ժ��ʽ", "��Ժ��ʽ", "����Ժ", "�Ƿ�ȷ��", "ȷ������", "ʬ���־", "�����־", "��������", _
                    "�·�����", "��ҽ�������", "���ȴ���", "�ɹ�����", "����ҽʦ", "סԺҽʦ", "����ҽʦ", "����ҽʦ", "���λ�ʿ", "����Ա���", "����Ա����")
            End If

            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = Get��ҳ����SQL(1, arrField) '��ȡZL_������ҳ_��ҳ����ĵ���SQL
        End If
    End If
End Function

Public Sub PopPatiAuxiSQL(ByRef arrSQL As Variant, Optional ByVal bln��ʿվ As Boolean)
'���ܣ���ȡһ��Ĵӱ���ϢSQL��������SQL������
'SaveMedPageData���Ӻ���
    Dim arrField As Variant, arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long, LngCol As Long
    Dim strTmp As String, arrTag As Variant

    arrField = Array()
    If bln��ʿվ Then
        gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
        strTmp = ",ѹ�������ڼ�,ѹ������,������׹���˺�,������׹��ԭ��,�����¼�,"
        If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
            strTmp = strTmp & "����Լ��,Լ����ʱ��,Լ����ʽ,Լ������,Լ��ԭ��,"
        ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
            strTmp = strTmp & "��Һҩ��,��Һ����,����Լ��,͸�����ص�ֵ,��Һ��Ӧ,"
        End If
        '��ʿվ�༭�Ĵӱ���Ϣ
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            gclsPros.MainInfoRec.MoveFirst
            For i = 1 To gclsPros.MainInfoRec.RecordCount
                If strTmp Like "*," & gclsPros.MainInfoRec!��Ϣ�� & ",*" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & gclsPros.MainInfoRec!��Ϣ�� & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "')"
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
        End If
    Else
        If gclsPros.PatiType = PF_סԺ Then
        '����������Ŀ�ı��棬��Ϊ����������Ŀ���ܻ���һ��Ĵӱ���Ϣ����ͬ�����������
        '��������������Ŀ������һ��ӱ���Ϣ������ͬ������һ��ӱ���ϢֵΪ׼����ˣ��ȱ��没��������Ŀ
        '����������Ŀ�ı���
            gclsPros.MainInfoRec.Filter = "��Ϣ��='������Ŀ' And �Ƿ�ı�=1"
            If gclsPros.MainInfoRec.RecordCount <> 0 Then
                gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬<>0 "
                gclsPros.SecdInfoRec.Sort = "Sort"
                With gclsPros.CurrentForm.vsfMain
                    For i = 1 To gclsPros.SecdInfoRec.RecordCount
                        arrTag = Split(gclsPros.SecdInfoRec!Tag & "", ";")
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & arrTag(0) & "','" & gclsPros.SecdInfoRec!��Ϣ��ֵ & "')"
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                End With
            End If
        End If

        gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
        If gclsPros.MainInfoRec.RecordCount <= 0 Then Exit Sub
        '3��һ��Ĳ�����ҳ�ӱ���Ϣ�ı���
        If gclsPros.PatiType = PF_���� Then
            '������ҳ�ӱ���Ϣ����
            '����������Ϣ����
            strTmp = gclsPros.CurrentForm.UCPatiVitalSigns.GetSaveSQL(gclsPros.����ID, gclsPros.��ҳID)
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strTmp
            End If
            gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
            arrField = Array("�Ļ��̶�", "����״��", "ȥ��", "RH", "Ѫ��", "ҽѧ��ʾ", "����ҽѧ��ʾ", "���֤��״̬", "�޹�����¼", "�⼮���֤��", "�໤�����֤��")
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.Sort = "���,��Ϣ��"
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!��Ϣ�� Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        If arrField(i) = "ҽѧ��ʾ" Or arrField(i) = "����ҽѧ��ʾ" Or arrField(i) = "RH" Or arrField(i) = "Ѫ��" Or arrField(i) = "���֤��״̬" Or arrField(i) = "�⼮���֤��" Then
                            arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & gclsPros.����ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "')"
                            If arrField(i) = "RH" Or arrField(i) = "Ѫ��" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & gclsPros.����ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "'," & gclsPros.��ҳID & ")"
                            End If
                        Else
                            arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & gclsPros.����ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "'," & gclsPros.��ҳID & ")"
                        End If
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
            Next
        Else
            arrField = Array("��Ժ����", "��Ժ����", "ת�Ƽ�¼", "��ҽΣ��", "��ҽ��֢", "��ҽ����", "��ҽ���ȷ���", _
                "������ҩ�Ƽ�", "��������ԭ��", "����ʱ��", "��Ժǰ����Ժ����", "ʾ�̲���", "���в���", "���Ѳ���", "RH", _
                "��Ѫ��Ӧ", "���ϸ��", "��ѪС��", "��Ѫ��", "��ȫѪ", "������", "������", "����ҽʦ", "����ҽʦ", "����ҽʦ", _
                IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, "����ҽʦ", "�о���ʵϰҽʦ"), "ʵϰҽʦ", "�ʿ�ҽʦ", "�ʿػ�ʿ", "��ԭѧ���", "��Ѫ���", "��ɫ������", "������", _
                "��������", "��Ⱦ����", "��Ժת��", "����Ժ�ƻ�����", "31������סԺ", "������������", "��������������", _
                "��������Ժ����", "������ʹ��ʱ��", "����ʱ��", "���Ȳ���", "�������", "�ֻ��̶�", "����������", _
                "��ҽ�豸", "��ҽ����", "��֤ʩ��", "�����", "��������", "��ҳ��������", "����״��", "����ʱ��", _
                "��Ⱦ��������ϵ", "��Ⱦ��λ", "����", "��Ժת��", "��Ժ��ʽ", "��ϵ�˸�����Ϣ", "���֤��״̬", "�޹�����¼", "סԺ�����ڼ�", "�⼮���֤��", "�໤�����֤��")
            'ҽѧ��ʾ���ڲ��˽���������������Ҫ
            strTmp = IIf(gclsPros.FuncType <> f������ҳ, "ҽѧ��ʾ,����ҽѧ��ʾ", "�ջ�����,ҽ����,��ҳX�ߺ�,�ؼ���������,һ����������,������������,������������,ICU����,CCU����,ת��ʱ��")
            '�Ĵ�����Һ�ɻ�ʿ��д,CT��MRI��Ϣ�������������У����жϵ�������ʱ���ٱ���
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, ",Ժ�ڻ���,��Ժ����,�������,��׵���,��������,����һ��ס��Ժʱ��,�Ƿ���ͬһ����", ",HBSAG,HCV-AB,HIV-AB,��Һ��Ӧ,CT,MRI")
            '�ٴ�·����Ϣ����׼������ϰ�û��
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_��������׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼, "", ",�ٴ�·��,�˳�ԭ��,����ԭ��,�没�ز�Σ")
            '���ϰ�����������ڵ���Ϣ
            strTmp = strTmp & IIf(gclsPros.MedPageSandard <> ST_����ʡ��׼, "", ",��֢�໤����,��֢�໤Сʱ,������,�ٴ�·��,��������,����T,����M,����N,�걾�ͼ�,��Ⱦ��,APGAR,DRGS")
            '���ϰ������������Ժ��ʽ
            strTmp = strTmp & IIf(gclsPros.MedPageSandard <> ST_����ʡ��׼, "", ",��������Ժ��ʽ,Χ��������,�������")
            '�������϶��У�����һ��ס��Ժʱ��,�Ƿ���ͬһ����
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_����ʡ��׼ And gclsPros.FuncType = f������ҳ, ",����һ��ס��Ժʱ��,�Ƿ���ͬһ����", "")
            arrTmp = Split(strTmp, ",")
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = Trim(arrTmp(i))
            Next
            
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.MoveFirst
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!��Ϣ�� Then
                        If arrField(i) = "������" Then
                            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬<>" & CS_δ�ı�
                            Do While Not gclsPros.SecdInfoRec.EOF
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                                    strTmp = "'������" & (Val(gclsPros.SecdInfoRec!IndexEx & "") + 4) & "','" & gclsPros.CurrentForm.vsTSJC.TextMatrix(gclsPros.SecdInfoRec!IndexEx, 1) & "'"
                                Else
                                    strTmp = "'" & decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), TR_CT, "CT", TR_PETCT, "PETCT", TR_˫ԴCT, "˫ԴCT", _
                                                TR_XƬ, "XƬ", TR_B��, "B��", TR_�����Ķ�ͼ, "�����Ķ�ͼ", TR_MRI, "MRI", TR_ͬλ�ؼ��, "ͬλ�ؼ��") & "','" & Mid(gclsPros.CurrentForm.vsTSJC.TextMatrix(Val(gclsPros.SecdInfoRec!IndexEx & ""), 1), 1, 1) & "'"
                                End If
                                arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & strTmp & ")"
                                gclsPros.SecdInfoRec.MoveNext
                            Loop
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "')"
                            If gclsPros.FuncType <> f������ҳ Then '������ҳ�������в�����Ϣ�ӱ�ı���
                                '������Ϣ�ӱ���Ϣ
                                If arrField(i) = "Ѫ��" Or arrField(i) = "RH" Or arrField(i) = "ҽѧ��ʾ" Or arrField(i) = "����ҽѧ��ʾ" Or arrField(i) = "��ϵ�˸�����Ϣ" Or arrField(i) = "���֤��״̬" Or arrField(i) = "�⼮���֤��" Or arrField(i) = "�໤�����֤��" Then
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & gclsPros.����ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!��Ϣ��ֵ & "')"
                                End If
                            End If
                        End If
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
            Next
        End If
    End If
End Sub

Public Sub PopPatiDiagSQL(ByRef arrSQL As Variant, ByVal datCur As Date)
'���ܣ������SQL����SQL������
'SaveMedPageData���Ӻ���
'��ҽ���, ��ҽ���,��ԭѧ���
    Dim k As Integer, LngRow As Long, j As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim lngID As Long
    Dim strDiagRowIDs As String, strDiagNames As String
    
    On Error GoTo errH
    
    If gclsPros.FuncType <> f������ҳ Then
        Call MsgDis(gclsPros.DiseaseIDs, gclsPros.DiagIDs)
    End If
    For k = 0 To 1
        gclsPros.MainInfoRec.Filter = "��Ϣ��='" & IIf(k = 0, "��ҽ���", "��ҽ���") & "' And �Ƿ�ı�=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            Set vsTmp = IIf(k = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY)
            With vsTmp
                'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
                gclsPros.SecdInfoRec.Filter = "(�ı�״̬=" & CS_ɾ���� & " And ���=" & gclsPros.MainInfoRec!��� & ") OR (�ı�״̬=" & CS_�滻�� & " And ���=" & gclsPros.MainInfoRec!��� & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
                Do While Not gclsPros.SecdInfoRec.EOF
                    strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                    gclsPros.SecdInfoRec.MoveNext
                Loop
                If strTmp <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    '����ϵͳ�洢�������ϵͳ���
                    arrSQL(UBound(arrSQL)) = "Zl" & IIf(gclsPros.FuncType = f������ҳ, 3, "") & "_������ϼ�¼_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & ",NULL,NUll,'" & Mid(strTmp, 2) & "')"
                End If
                '����Ϣ�ı��Լ���������Ҫ���ò������
                '�μ���Ϣ�ı䣬���ø��¹���
                gclsPros.SecdInfoRec.Filter = "�ı�״̬>" & CS_δ�ı� & " And ���=" & gclsPros.MainInfoRec!���: gclsPros.SecdInfoRec.Sort = "Sort"
                Do While Not gclsPros.SecdInfoRec.EOF
                    LngRow = gclsPros.SecdInfoRec!IndexEx: j = Val(Mid(gclsPros.SecdInfoRec!��Ϣ��ֵ, 1, InStr(gclsPros.SecdInfoRec!��Ϣ��ֵ, "|") - 1))
                    If Trim(.TextMatrix(LngRow, DI_��ϱ���)) = "" Then
                        strTmp = .TextMatrix(LngRow, DI_�������) & IIf(.TextMatrix(LngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(LngRow, DI_��ҽ֤��) & ")", "")
                    Else
                        strTmp = "(" & .TextMatrix(LngRow, DI_��ϱ���) & ")" & .TextMatrix(LngRow, DI_�������) & IIf(.TextMatrix(LngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(LngRow, DI_��ҽ֤��) & ")", "")
                    End If

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!�ı�״̬ <> CS_������ Then
                        If gclsPros.FuncType = f������ҳ Then
                            arrSQL(UBound(arrSQL)) = "ZL3_������ϼ�¼_INSERT(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & .TextMatrix(LngRow, DI_��Ϸ���) & "," & _
                                    ZVal(.TextMatrix(LngRow, DI_����ID)) & "," & ZVal(.TextMatrix(LngRow, DI_���ID)) & "," & ZVal(.TextMatrix(LngRow, DI_֤��ID)) & ",'" & _
                                    strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_��Ժ���)) & "'," & IIf(.TextMatrix(LngRow, DI_�Ƿ�δ��) = "", 0, 1) & "," & _
                                    IIf(.TextMatrix(LngRow, DI_�Ƿ�����) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_��ע) & "','" & _
                                    .TextMatrix(LngRow, DI_��Ժ����) & "'," & ZVal(.TextMatrix(LngRow, DI_����ID)) & ")"
                            gclsPros.AddDiag = True
                        Else
                            lngID = zlDatabase.GetNextId("������ϼ�¼")
                            gclsPros.SecdInfoRec.Update "Tag", lngID '������ID
                            .RowData(LngRow) = lngID
                            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",3,NULL," & .TextMatrix(LngRow, DI_��Ϸ���) & "," & _
                                                ZVal(.TextMatrix(LngRow, DI_����ID)) & "," & ZVal(.TextMatrix(LngRow, DI_���ID)) & "," & ZVal(.TextMatrix(LngRow, DI_֤��ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_��Ժ���)) & "'," & IIf(.TextMatrix(LngRow, DI_�Ƿ�δ��) = "", 0, 1) & "," & _
                                                IIf(.TextMatrix(LngRow, DI_�Ƿ�����) = "", 0, 1) & "," & zlStr.To_Date(datCur, "ymdhms") & ",'" & .TextMatrix(LngRow, DI_ҽ��IDs) & "' ," & j & ",'" & .TextMatrix(LngRow, DI_��ע) & "','" & _
                                                .TextMatrix(LngRow, DI_��Ժ����) & "'," & zlStr.To_Date(.TextMatrix(LngRow, DI_����ʱ��), "ymdhm") & ",Null," & lngID & "," & ZVal(.TextMatrix(LngRow, DI_����ID)) & ")"
                            gclsPros.AddDiag = True
                        End If
                    Else
                        If gclsPros.FuncType = f������ҳ Then
                            arrSQL(UBound(arrSQL)) = "Zl3_������ϼ�¼_Update(" & gclsPros.SecdInfoRec!ID & "," & .TextMatrix(LngRow, DI_��Ϸ���) & "," _
                                                & ZVal(.TextMatrix(LngRow, DI_����ID)) & "," & ZVal(.TextMatrix(LngRow, DI_���ID)) & "," & ZVal(.TextMatrix(LngRow, DI_֤��ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_��Ժ���)) & "'," & IIf(.TextMatrix(LngRow, DI_�Ƿ�δ��) = "", 0, 1) & "," _
                                                & IIf(.TextMatrix(LngRow, DI_�Ƿ�����) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_��ע) & "','" & _
                                                .TextMatrix(LngRow, DI_��Ժ����) & "'," & ZVal(.TextMatrix(LngRow, DI_����ID)) & ")"
                        Else
                            arrSQL(UBound(arrSQL)) = "Zl_������ϼ�¼_Update(" & gclsPros.SecdInfoRec!ID & "," & gclsPros.����ID & "," & gclsPros.��ҳID & ",3," & .TextMatrix(LngRow, DI_��Ϸ���) & "," _
                                                & ZVal(.TextMatrix(LngRow, DI_����ID)) & "," & ZVal(.TextMatrix(LngRow, DI_���ID)) & "," & ZVal(.TextMatrix(LngRow, DI_֤��ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_��Ժ���)) & "'," & IIf(.TextMatrix(LngRow, DI_�Ƿ�δ��) = "", 0, 1) & "," _
                                                & IIf(.TextMatrix(LngRow, DI_�Ƿ�����) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_��ע) & "','" & _
                                                .TextMatrix(LngRow, DI_��Ժ����) & "'," & zlStr.To_Date(.TextMatrix(LngRow, DI_����ʱ��), "ymdhm") & "," & ZVal(.TextMatrix(LngRow, DI_����ID)) & ")"
                        End If
                    End If
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            End With
        End If
    Next
    
    '���ѡ����֯����ֵ
    If gclsPros.FuncType = f���ѡ�� Then
        For k = 0 To 1
            Set vsTmp = IIf(k = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY)
            With vsTmp
                For LngRow = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(LngRow, DI_����)) <> 0 Then
                        If Trim(.TextMatrix(LngRow, DI_��ϱ���)) = "" Then
                            strTmp = .TextMatrix(LngRow, DI_�������) & IIf(.TextMatrix(LngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(LngRow, DI_��ҽ֤��) & ")", "")
                        Else
                            strTmp = "(" & .TextMatrix(LngRow, DI_��ϱ���) & ")" & .TextMatrix(LngRow, DI_�������) & IIf(.TextMatrix(LngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(LngRow, DI_��ҽ֤��) & ")", "")
                        End If
                        strDiagRowIDs = strDiagRowIDs & "," & Val(.RowData(LngRow))
                        strDiagNames = strDiagNames & "," & strTmp
                    End If
                Next
            End With
        Next
        gclsPros.DiagRowIDs = Mid(strDiagRowIDs, 2)
        gclsPros.DiagNames = Mid(strDiagNames, 2)
    End If

    If gclsPros.PatiType = PF_סԺ And gclsPros.FuncType <> f���ѡ�� Then
        gclsPros.MainInfoRec.Filter = "��Ϣ��='��ԭѧ���' And �Ƿ�ı�=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & ",NULL,'21')"
            If Not NVL(gclsPros.MainInfoRec!��Ϣ��ֵ) = "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If gclsPros.FuncType = f������ҳ Then
                    arrSQL(UBound(arrSQL)) = "ZL3_������ϼ�¼_INSERT(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",21," & _
                        ZVal(gclsPros.CurrentForm.cmdInfo(GC_��ԭѧ���).Tag) & ",NULL,NULL,'" & gclsPros.CurrentForm.txtInfo(GC_��ԭѧ���).Text & "',NULL,NULL,NULL)"
                Else
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",3,NULL,21," & _
                            ZVal(gclsPros.CurrentForm.cmdInfo(GC_��ԭѧ���).Tag) & ",NULL,NULL,'" & gclsPros.CurrentForm.txtInfo(GC_��ԭѧ���).Text & "',NULL,NULL,NULL," & _
                            zlStr.To_Date(datCur, "ymdhms") & ",Null,1,Null,Null)"
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub PopKSSSQL(ByRef arrSQL As Variant)
'���ܣ���������SQL��������
'SaveMedPageData���Ӻ���
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant

     'ʹ�ÿ����صļ�¼
    gclsPros.MainInfoRec.Filter = "��Ϣ��='���˿����ؼ�¼' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsKSS
        With vsTmp
            gclsPros.SecdInfoRec.Filter = "�ı�״̬<>" & CS_δ�ı� & " And ���=" & gclsPros.MainInfoRec!���: gclsPros.SecdInfoRec.Sort = "�ı�״̬,Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                'ɾ�����Լ�����Ϣ�ı�����Ҫ���ù��̴��빦��2��ɾ����
                If gclsPros.SecdInfoRec!�ı�״̬ = CS_ɾ���� Or gclsPros.SecdInfoRec!�ı�״̬ = CS_�滻�� Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrTmp = Split(gclsPros.SecdInfoRec!��Ϣԭֵ, "|")
                    arrSQL(UBound(arrSQL)) = "Zl_���˿����ؼ�¼_Update(" & _
                            "2," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & Val(arrTmp(0)) & ",'" & arrTmp(1) & "','" & Trim(arrTmp(2)) & "','" & Trim(arrTmp(3)) & "')"
                End If
                '����Ϣ�ı��Լ���������Ҫ���ù��̴��빦��0��������
                '�μ���Ϣ�ı䣬���ù��̴��빦��1���޸ģ�
                If gclsPros.SecdInfoRec!�ı�״̬ > CS_δ�ı� Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    LngRow = gclsPros.SecdInfoRec!IndexEx
                    arrSQL(UBound(arrSQL)) = "Zl_���˿����ؼ�¼_Update(" & _
                                decode(gclsPros.SecdInfoRec!�ı�״̬, 1, 1, 2, 0, 3, 0) & "," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & Val(.RowData(LngRow) & "") & ",'" & .TextMatrix(LngRow, KI_����ҩ����) & "','" & _
                                Trim(.TextMatrix(LngRow, KI_��ҩĿ��)) & "','" & Trim(.TextMatrix(LngRow, KI_ʹ�ý׶�)) & "'," & Val(.TextMatrix(LngRow, KI_ʹ������)) & ",'" & UserInfo.���� & "',Sysdate," & _
                                ZVal(.Cell(flexcpChecked, LngRow, KI_һ���п�Ԥ����)) & "," & ZVal(.TextMatrix(LngRow, KI_DDD��)) & ",'" & .TextMatrix(LngRow, KI_������ҩ) & "')"
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
        '��������-������ҳ�ӱ������һ��ɾȥ
        gclsPros.AuxiInfo.Filter = "��Ϣ�� Like '������*'": gclsPros.AuxiInfo.Sort = "��Ϣ��"
        If gclsPros.AuxiInfo.RecordCount > 0 Then
            Do While Not gclsPros.AuxiInfo.EOF
                If IsNumeric(Mid(gclsPros.AuxiInfo!��Ϣ��, 4)) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & gclsPros.AuxiInfo!��Ϣ�� & "',NULL)"
                End If
                gclsPros.AuxiInfo.MoveNext
            Loop
        End If
    End If

End Sub

Private Sub PopOPSSQL(ByRef arrSQL As Variant)
'���ܣ�������SQL��������
'SaveMedPageData���Ӻ���
    Dim LngRow As Long, lngOrder As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant
    Dim lngID As Long
    Dim OPSid As Long
    
      '�������
    gclsPros.MainInfoRec.Filter = "��Ϣ��='�������' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsOPS
        With vsTmp
            'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
            gclsPros.SecdInfoRec.Filter = "(�ı�״̬=" & CS_ɾ���� & " And ���=" & gclsPros.MainInfoRec!��� & ") OR (�ı�״̬=" & CS_�滻�� & " And ���=" & gclsPros.MainInfoRec!��� & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
            Do While Not gclsPros.SecdInfoRec.EOF
                strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                gclsPros.SecdInfoRec.MoveNext
            Loop
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���������¼_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & ",Null,'" & Mid(strTmp, 2) & "')"
            End If
            
            '����Ϣ�ı��Լ���������Ҫ���ò������
            '�μ���Ϣ�ı䣬���ø��¹���
            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬>" & CS_δ�ı�: gclsPros.SecdInfoRec.Sort = "Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                LngRow = gclsPros.SecdInfoRec!IndexEx: lngOrder = GetOPSOrder(vsTmp, LngRow)
                strTmp = Trim(.TextMatrix(LngRow, PI_�п�����))
                If strTmp = "" Then strTmp = "/"
                strTmp = strTmp & "/" & decode(.TextMatrix(LngRow, PI_��������), "һ������", 1, "��������", 2, "��������", 3, "�ļ�����", 4, "��", 9, 0)
                arrTmp = Split(strTmp, "/")
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                '�����������������ڽ������в���ʾʱ���������ھ͵�����������
                If gclsPros.MedPageSandard = ST_��������׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    If Not gclsPros.UseOPSEndTime Then
                        .TextMatrix(LngRow, PI_��������) = .TextMatrix(LngRow, PI_��������)
                    End If
                ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    .TextMatrix(LngRow, PI_��������) = .TextMatrix(LngRow, PI_��������)
                End If

                If gclsPros.SecdInfoRec!�ı�״̬ <> CS_������ Then
                    lngID = zlDatabase.GetNextId("���������¼")
                    arrSQL(UBound(arrSQL)) = "ZL_���������¼_Insert(" & lngID & "," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "," & lngOrder & "," & _
                            zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & _
                            "NULL," & ZVal(.TextMatrix(LngRow, PI_��������ID)) & "," & ZVal(.TextMatrix(LngRow, PI_������ĿID)) & ",'" & .TextMatrix(LngRow, PI_��������) & "','" & .TextMatrix(LngRow, PI_����ҽʦ) & "','" & _
                            .TextMatrix(LngRow, PI_������ʿ) & "','" & .TextMatrix(LngRow, PI_����1) & "','" & .TextMatrix(LngRow, PI_����2) & "',NULL," & zlStr.To_Date(.TextMatrix(LngRow, PI_����ʼʱ��), "ymdhm") & ",NULL," & ZVal(.TextMatrix(LngRow, PI_����ID)) & ",'" & _
                            .TextMatrix(LngRow, PI_��������) & "',NULL,NULL,'" & .TextMatrix(LngRow, PI_����ҽʦ) & "',NULL,NULL,'" & arrTmp(0) & "','" & arrTmp(1) & "',Sysdate,'" & .TextMatrix(LngRow, PI_�������) & "','" & _
                            .TextMatrix(LngRow, PI_ASA�ּ�) & "'," & Abs(Val(.TextMatrix(LngRow, PI_�ٴ�����))) & ",'" & .TextMatrix(LngRow, PI_NNIS�ּ�) & "'," & arrTmp(2) & "," & ZVal(.TextMatrix(LngRow, PI_׼������)) & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_������ҩʱ��), "ymdhm") & ",'" & _
                            .TextMatrix(LngRow, PI_�пڲ�λ) & "'," & .Cell(flexcpChecked, LngRow, PI_�пڸ�Ⱦ) & "," & .Cell(flexcpChecked, LngRow, PI_����֢) & "," & .Cell(flexcpChecked, LngRow, PI_�ط������Ҽƻ�) & ",'" & .TextMatrix(LngRow, PI_�ط�������Ŀ��) & "'," & _
                            .Cell(flexcpChecked, LngRow, PI_Ԥ���ÿ���ҩ) & "," & Val(.TextMatrix(LngRow, PI_����ҩ����)) & "," & .Cell(flexcpChecked, LngRow, PI_��Ԥ�ڵĶ�������) & "," & .Cell(flexcpChecked, LngRow, PI_������֢) & "," & .Cell(flexcpChecked, LngRow, PI_������������) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_��������֢) & "," & .Cell(flexcpChecked, LngRow, PI_�����Ѫ��Ѫ��) & "," & .Cell(flexcpChecked, LngRow, PI_�����˿��ѿ�) & "," & .Cell(flexcpChecked, LngRow, PI_�������Ѫ˨) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_���������л����) & "," & .Cell(flexcpChecked, LngRow, PI_�������˥��) & "," & .Cell(flexcpChecked, LngRow, PI_�����˨��) & "," & .Cell(flexcpChecked, LngRow, PI_�����Ѫ֢) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_�����Źؽڹ���) & ")"
                    .RowData(LngRow) = lngID
                    gclsPros.SecdInfoRec!ID = lngID
                Else
                    OPSid = IIf(.RowData(LngRow) <> "", .RowData(LngRow), Val(gclsPros.SecdInfoRec!ID & ""))
                    arrSQL(UBound(arrSQL)) = "ZL_���������¼_Update(" & OPSid & "," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "," & lngOrder & "," & _
                            zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_��������), "ymdhm") & "," & _
                            "NULL," & ZVal(.TextMatrix(LngRow, PI_��������ID)) & "," & ZVal(.TextMatrix(LngRow, PI_������ĿID)) & ",'" & .TextMatrix(LngRow, PI_��������) & "','" & .TextMatrix(LngRow, PI_����ҽʦ) & "','" & _
                            .TextMatrix(LngRow, PI_������ʿ) & "','" & .TextMatrix(LngRow, PI_����1) & "','" & .TextMatrix(LngRow, PI_����2) & "',NULL," & zlStr.To_Date(.TextMatrix(LngRow, PI_����ʼʱ��), "ymdhm") & ",NULL," & ZVal(.TextMatrix(LngRow, PI_����ID)) & ",'" & _
                            .TextMatrix(LngRow, PI_��������) & "',NULL,NULL,'" & .TextMatrix(LngRow, PI_����ҽʦ) & "',NULL,NULL,'" & arrTmp(0) & "','" & arrTmp(1) & "','" & .TextMatrix(LngRow, PI_�������) & "','" & _
                            .TextMatrix(LngRow, PI_ASA�ּ�) & "'," & Abs(Val(.TextMatrix(LngRow, PI_�ٴ�����))) & ",'" & .TextMatrix(LngRow, PI_NNIS�ּ�) & "'," & arrTmp(2) & "," & ZVal(.TextMatrix(LngRow, PI_׼������)) & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_������ҩʱ��), "ymdhm") & ",'" & _
                            .TextMatrix(LngRow, PI_�пڲ�λ) & "'," & .Cell(flexcpChecked, LngRow, PI_�пڸ�Ⱦ) & "," & .Cell(flexcpChecked, LngRow, PI_����֢) & "," & .Cell(flexcpChecked, LngRow, PI_�ط������Ҽƻ�) & ",'" & .TextMatrix(LngRow, PI_�ط�������Ŀ��) & "'," & _
                            .Cell(flexcpChecked, LngRow, PI_Ԥ���ÿ���ҩ) & "," & Val(.TextMatrix(LngRow, PI_����ҩ����)) & "," & .Cell(flexcpChecked, LngRow, PI_��Ԥ�ڵĶ�������) & "," & .Cell(flexcpChecked, LngRow, PI_������֢) & "," & .Cell(flexcpChecked, LngRow, PI_������������) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_��������֢) & "," & .Cell(flexcpChecked, LngRow, PI_�����Ѫ��Ѫ��) & "," & .Cell(flexcpChecked, LngRow, PI_�����˿��ѿ�) & "," & .Cell(flexcpChecked, LngRow, PI_�������Ѫ˨) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_���������л����) & "," & .Cell(flexcpChecked, LngRow, PI_�������˥��) & "," & .Cell(flexcpChecked, LngRow, PI_�����˨��) & "," & .Cell(flexcpChecked, LngRow, PI_�����Ѫ֢) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_�����Źؽڹ���) & ")"
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
    End If
End Sub

Private Function GetOPSOrder(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long) As Long
'���ܣ���ȡָ���������¼�Ĵ���
    Dim i As Long, lngOrder As Long
    
    With vsOPS
        For i = .FixedRows To LngRow
            If .TextMatrix(i, PI_��������) <> "" Then
                lngOrder = lngOrder + 1
            End If
        Next
    End With
    GetOPSOrder = lngOrder
End Function

Private Sub PopAllerSQL(ByRef arrSQL As Variant)
'���ܣ���������ϢSQL��������
'SaveMedPageData���Ӻ���
    Dim k As Integer, LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    
    '������Ϣ
    gclsPros.MainInfoRec.Filter = "��Ϣ��='����ҩ��' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsAller
        With vsTmp
            'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
            gclsPros.SecdInfoRec.Filter = "(�ı�״̬=" & CS_ɾ���� & " And ���=" & gclsPros.MainInfoRec!��� & ") OR (�ı�״̬=" & CS_�滻�� & " And ���=" & gclsPros.MainInfoRec!��� & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
            Do While Not gclsPros.SecdInfoRec.EOF
                strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                gclsPros.SecdInfoRec.MoveNext
            Loop
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���˹�����¼_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & ",'" & Mid(strTmp, 2) & "')"
            End If
            '����Ϣ�ı��Լ���������Ҫ���ò������
            '�μ���Ϣ�ı䣬���ø��¹���
            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬>" & CS_δ�ı�: gclsPros.SecdInfoRec.Sort = "Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                LngRow = gclsPros.SecdInfoRec!IndexEx
                If .TextMatrix(LngRow, AI_����ҩ��) <> "��" Then
    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!�ı�״̬ <> CS_������ Then
                        arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & _
                                IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "," & ZVal(.TextMatrix(LngRow, AI_ҩ��ID)) & ",'" & .TextMatrix(LngRow, AI_����ҩ��) & "',1," & _
                                zlStr.To_Date(.TextMatrix(LngRow, AI_����ʱ��), "ymd") & ",SysDate,'" & _
                                .TextMatrix(LngRow, AI_������Ӧ) & "','" & .TextMatrix(LngRow, AI_����Դ����) & "')"
                        gclsPros.AddAller = True
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_���˹�����¼_Update(" & gclsPros.SecdInfoRec!ID & "," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & _
                                IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "," & ZVal(.TextMatrix(LngRow, AI_ҩ��ID)) & ",'" & .TextMatrix(LngRow, AI_����ҩ��) & "',1," & _
                                zlStr.To_Date(.TextMatrix(LngRow, AI_����ʱ��), "ymd") & ",'" & _
                                .TextMatrix(LngRow, AI_������Ӧ) & "','" & .TextMatrix(LngRow, AI_����Դ����) & "')"
                    End If
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub PopICUSQL(ByRef arrSQL As Variant)
'���ܣ�����֢�໤��������SQL��������
'SaveMedPageData���Ӻ���
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrFields As Variant
    
    If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
        gclsPros.MainInfoRec.Filter = "��Ϣ��='�໤������' OR ��Ϣ��='�˹������ѳ�' OR ��Ϣ��='�ط���֢ҽѧ��' OR ��Ϣ��='�ط����ʱ��'"
        Set rsTmp = Rec.FilterNew(gclsPros.MainInfoRec)
        rsTmp.Filter = "�Ƿ�ı�=1"
        If rsTmp.RecordCount > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������֢�໤���_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������֢�໤���_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",1"
            arrFields = Array("�໤������", "", "", "", "", "�˹������ѳ�", "�ط���֢ҽѧ��", "�ط����ʱ��")
            For LngRow = LBound(arrFields) To UBound(arrFields)
                If arrFields(LngRow) = "" Then
                    arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ",Null"
                Else
                    rsTmp.Filter = "��Ϣ��='" & arrFields(LngRow) & "'"
                    If arrFields(LngRow) = "�໤������" Or arrFields(LngRow) = "�ط����ʱ��" Then
                        arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ",'" & rsTmp!��Ϣ��ֵ & "'"
                    Else
                         arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & "," & ZVal(rsTmp!��Ϣ��ֵ & "")
                    End If
                End If
            Next
            arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ")"
        End If
    Else
        '��֢�໤��¼
        gclsPros.MainInfoRec.Filter = "��Ϣ��='������֢�໤���' And �Ƿ�ı�=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
            With vsTmp
                '��˴��������ֻ�м���ɾ��
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������֢�໤���_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
                For LngRow = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(LngRow, UI_�໤������)) <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_������֢�໤���_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & LngRow & ",'" & Trim(.TextMatrix(LngRow, UI_�໤������)) & "'," & _
                                    zlStr.To_Date(.TextMatrix(LngRow, UI_����ʱ��), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, UI_�˳�ʱ��), "ymdhm") & "," & ZVal(.TextMatrix(LngRow, UI_����ס�ƻ�)) & ",'" & .TextMatrix(LngRow, UI_����סԭ��) & "')"
                    End If
                Next
            End With
        End If
    End If
End Sub

Public Sub PopDiagMatchSQL(ByRef arrSQL As Variant)
'���ܣ�����Ϸ��������SQL��������
'SaveMedPageData���Ӻ���
    Dim arrField As Variant, arrFieldEx As Variant
    Dim i As Long
    
     '��Ϸ������
    gclsPros.MainInfoRec.Filter = "��Ϣ��='��Ϸ������' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        gclsPros.SecdInfoRec.Filter = "�ı�״̬<>" & CS_δ�ı� & " And ���=" & gclsPros.MainInfoRec!���: gclsPros.SecdInfoRec.Sort = "Sort"
        If gclsPros.SecdInfoRec.RecordCount > 0 Then
            arrField = Array(BCC_�������ԺXY, BCC_��Ժ���ԺXY, BCC_�����벡��, BCC_�ٴ��벡��, BCC_�ٴ���ʬ��, BCC_��ǰ������, BCC_��������Ժ, _
                    BCC_�������ԺZY, BCC_��Ժ���ԺZY, BCC_��֤, BCC_�η�, BCC_��ҩ)
            arrFieldEx = Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15)
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.SecdInfoRec.MoveFirst
                Do While Not gclsPros.SecdInfoRec.EOF
                    If arrField(i) = gclsPros.SecdInfoRec!IndexEx Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & arrFieldEx(i) & "," & IIf(gclsPros.SecdInfoRec!��Ϣ��ֵ & "" = "", "Null", gclsPros.SecdInfoRec!��Ϣ��ֵ) & ")"
                    End If
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            Next
        End If
    End If
End Sub

Private Sub PopStructAdressSQL(ByRef arrSQL As Variant)
'���ܣ����ṹ����ַ��SQL��������
'SaveMedPageData���Ӻ���
    Dim arrField As Variant
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngType As Long
    
    With gclsPros.CurrentForm
        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
            arrField = Array("�����ص�", "����", "��ͥ��ַ", "���ڵ�ַ", "��ϵ�˵�ַ", "��λ��ַ")
        Else
            arrField = Array("�����ص�", "����", "��ͥ��ַ", "���ڵ�ַ", "��ϵ�˵�ַ")
        End If
        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.Filter = "��Ϣ��='" & arrField(i) & "'"
            blnAdd = False
            If Not gclsPros.MainInfoRec.EOF Then
                strTmp = .padrInfo(gclsPros.MainInfoRec!Index).valueʡ & "," & .padrInfo(gclsPros.MainInfoRec!Index).value�� & "," & .padrInfo(gclsPros.MainInfoRec!Index).value����
                If .padrInfo(gclsPros.MainInfoRec!Index).Items > 3 Then
                    strTmp = strTmp & "," & IIf(.padrInfo(gclsPros.MainInfoRec!Index).Items = 4, "," & .padrInfo(gclsPros.MainInfoRec!Index).value����, .padrInfo(gclsPros.MainInfoRec!Index).value���� & ",") & .padrInfo(gclsPros.MainInfoRec!Index).value��ϸ��ַ & "," & .padrInfo(gclsPros.MainInfoRec!Index).Code
                Else
                    strTmp = strTmp & ",,," & .padrInfo(gclsPros.MainInfoRec!Index).Code
                End If
                If Trim(Replace(strTmp, ",", "")) = "" Then
                    strTmp = ""
                Else
                    strTmp = Replace(strTmp, ",", "','")
                End If
                If gclsPros.MainInfoRec!�Ƿ�ı� = 0 Then 'δ�����ı䣬��û�нṹ��ַ��Ϣ�����������
                    Set rsTmp = GetStrucAddress(gclsPros.����ID, gclsPros.��ҳID, arrField(i))
                    blnAdd = rsTmp.EOF
                Else
                    blnAdd = True
                End If
                lngType = IIf(strTmp = "", 2, 1)
                If blnAdd Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(" & lngType & "," & gclsPros.����ID & "," & gclsPros.��ҳID & "," & (i + 1) & ",'" & strTmp & "')"
                End If
            End If
        Next
    End With
End Sub

Private Sub PopFeeSQL(ByRef arrSQL As Variant)
'���ܣ������˷��ñ�������SQL��������
'SaveMedPageData���Ӻ���
    Dim i As Long, LngRow As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid

    gclsPros.MainInfoRec.Filter = "��Ϣ��='���˷���' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_���˷���_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
        Set vsTmp = gclsPros.CurrentForm.vsFees
        With vsTmp
            For i = .FixedRows * 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) <> "" Then '�������ǿ�
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_���˷���_insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & _
                        ",'" & GetTextByDot(.TextMatrix(LngRow, LngCol)) & "'," & Val(.TextMatrix(LngRow, LngCol + 1)) & ")"
                End If
            Next
        End With
    End If
End Sub

Public Sub PopDelicerySQL(ByRef arrSQL As Variant)
'���ܣ���������Ϣ��SQL��������
'SaveMedPageData���Ӻ���
'   1-SaveMedPageData���Ӻ���
'   2-�������Ǽ�ʱ���������Ϣ¼��
        If grsDeliceryInfo Is Nothing Then Exit Sub
        '�������ӱ���Ϣ
        grsDeliceryInfo.Filter = "����=0 And ��¼����=1": grsDeliceryInfo.Sort = "��Ϣ��"
        Do While Not grsDeliceryInfo.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & grsDeliceryInfo!��Ϣ�� & "','" & grsDeliceryInfo!��Ϣ��ֵ & "')"
            grsDeliceryInfo.MoveNext
        Loop
        grsBabyInfo.Filter = "��¼����=1"
        grsBabyDiag.Filter = "��¼����=1"
        '���������������Ϣ
        If Not grsBabyInfo.EOF Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",0)"
        ElseIf Not grsBabyDiag.EOF Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",1)"
        End If

        Do While Not grsBabyInfo.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_���˷�����Ϣ_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & grsBabyInfo!̥������ & ",'" & grsBabyInfo!���䷽ʽ & "','" & _
                        grsBabyInfo!����̥λ & "','" & grsBabyInfo!������� & "', '" & IIf((grsBabyInfo!����ȱ�� & "") = 0 And grsBabyInfo!����ȱ�� & "" <> "��", 0, 1) & "', '" & grsBabyInfo!Ӥ���Ա� & "','" & grsBabyInfo!Ӥ������ & "', '" & grsBabyInfo!Apgar���� & "',to_Date('" & grsBabyInfo!����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ")"
            grsBabyInfo.MoveNext
        Loop
        Do While Not grsBabyDiag.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_��������ϼ�¼_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & grsBabyDiag!̥������ & "," & grsBabyDiag!��ϴ��� & ", " & grsBabyDiag!����id & ",'" & grsBabyDiag!������Ϣ & "')"
            grsBabyDiag.MoveNext
        Loop
End Sub

Private Sub PopShareInfoSQL(ByRef arrSQL As Variant)
'���ܣ������ƣ����ƣ�����ҩƷ��Ϣ��SQL�������飬������Ϣ���ڲ���ϵͳ��סԺ��ҳ�ڲ�������ʱ�ű���
'SaveMedPageData���Ӻ���
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long
    Dim vsTmp As VSFlexGrid
    
    strTmp = "������������,�������Ƽ�¼,�������Ƽ�¼"
    arrTmp = Split(strTmp, ",")
    strTmp = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
        gclsPros.MainInfoRec.Filter = "��Ϣ��='" & arrTmp(i) & "' And �Ƿ�ı�=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            If arrTmp(i) = "������������" Then
                Set vsTmp = gclsPros.CurrentForm.vsSpirit
            ElseIf arrTmp(i) = "�������Ƽ�¼" Then
                Set vsTmp = gclsPros.CurrentForm.vsChemoth
            Else
                Set vsTmp = gclsPros.CurrentForm.vsRadioth
            End If
            With vsTmp
                'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
                gclsPros.SecdInfoRec.Filter = "(�ı�״̬=" & CS_ɾ���� & " And ���=" & gclsPros.MainInfoRec!��� & ") OR (�ı�״̬=" & CS_�滻�� & " And ���=" & gclsPros.MainInfoRec!��� & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
                Do While Not gclsPros.SecdInfoRec.EOF
                    strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                    gclsPros.SecdInfoRec.MoveNext
                Loop
                If strTmp <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & Mid(strTmp, 2) & "')"
                End If
                '����Ϣ�ı��Լ���������Ҫ���ò������
                '�μ���Ϣ�ı䣬���ø��¹���
                gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬>" & CS_δ�ı�: gclsPros.SecdInfoRec.Sort = "Sort"
                Do While Not gclsPros.SecdInfoRec.EOF
                    LngRow = gclsPros.SecdInfoRec!IndexEx
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!�ı�״̬ <> CS_������ Then
                        arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & LngRow
                        gclsPros.SecdInfoRec!ID = LngRow
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Update(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & LngRow
                    End If
                    For j = .FixedCols To .Cols - 2 '��һ���������-2
                        Select Case j
                            Case RI_�������Ʊ��� 'CI_��ѧ���Ʊ���, SI_ҩ������
                                If i = 0 Then
                                    strTmp = ZVal(.TextMatrix(LngRow, SI_ҩƷID)) & ",'" & .TextMatrix(LngRow, j) & "'"
                                Else
                                    strTmp = ZVal(.TextMatrix(LngRow, IIf(i = 1, CI_����ID, RI_����ID)))
                                End If
                            Case RI_��ʼ����, RI_�������� 'CI_��ʼ����, CI_��������; SI_�Ƴ� ,SI_�������
                                If i = 0 Then '����ҩƷ
                                    strTmp = "'" & .TextMatrix(LngRow, j) & "'"
                                Else '���ƻ���
                                    strTmp = zlStr.To_Date(.TextMatrix(LngRow, j), "ymd")
                                End If
                            Case RI_��Ұ��λ, RI_������� 'CI_�Ƴ���,SI_���ⷴӦ;CI_���Ʒ���, SI_��Ч
                                '���Ʒ�������뻯���Ƴ�����Ϊ������
                                strTmp = IIf(i = 2 And j = RI_������� Or i = 1 And j = CI_�Ƴ���, Val(.TextMatrix(LngRow, j)), "'" & .TextMatrix(LngRow, j) & "'")
                            Case RI_�ۼ��� 'CI_����
                                strTmp = Val(.TextMatrix(LngRow, j)) & ""
                            Case Else
                                strTmp = "'" & .TextMatrix(LngRow, j) & "'"
                        End Select
                        arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & "," & strTmp
                    Next
                    arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ")"
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            End With
        End If
    Next
End Sub
Private Sub PopOtherSQL(ByRef arrSQL As Variant)
'���ܣ�����֢��е��ҽԺ��Ⱦ���걾�����������
'SaveMedPageData���Ӻ���
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant
    '��֢��е����ʹ�����
    gclsPros.MainInfoRec.Filter = "��Ϣ��='��е����ʹ�����' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
        With vsTmp
            '��Ϊ���������ֻ�м���ɾ��
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_��е����ʹ�����_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, TI_��е������)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_��е����ʹ�����_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & Val(.Cell(flexcpData, LngRow, TI_ICU����)) & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, TI_ICU����)), , "-") & "','" & GetTextByDot(Trim(.TextMatrix(LngRow, TI_��е������)), True) & "'," & _
                                                zlStr.To_Date(.TextMatrix(LngRow, TI_��ʼʱ��), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, TI_����ʱ��), "ymdhm") & ",'" & .TextMatrix(LngRow, TI_��Ⱦ�ۼ�Сʱ) & "')"
                End If
            Next
        End With
    End If
    '���˸�Ⱦ��¼
    gclsPros.MainInfoRec.Filter = "��Ϣ��='���˸�Ⱦ��¼' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsInfect
        With vsTmp
            '���������Ϊ����������ȡɾ�����������������Ż���
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
             arrSQL(UBound(arrSQL)) = "Zl_���˸�Ⱦ��¼_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, FI_��Ⱦ��λ)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_���˸�Ⱦ��¼_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & LngRow & "," & zlStr.To_Date(.TextMatrix(LngRow, FI_ȷ������), "ymdhm") & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, FI_��Ⱦ��λ))) & "','" & .TextMatrix(LngRow, FI_ҽԺ��Ⱦ����) & "')"
                End If
            Next
        End With
    End If
    
    '���˲�ԭѧ���
    gclsPros.MainInfoRec.Filter = "��Ϣ��='���˲�ԭѧ���' And �Ƿ�ı�=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsSample
        With vsTmp
            '���������Ϊ����������ȡɾ�����������������Ż���
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
             arrSQL(UBound(arrSQL)) = "Zl_���˲�ԭѧ���_Delete(" & gclsPros.����ID & "," & gclsPros.��ҳID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, MI_�걾)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_���˲�ԭѧ���_Insert(" & gclsPros.����ID & "," & gclsPros.��ҳID & "," & LngRow & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, MI_�걾)), True) & "','" & GetTextByDot(.TextMatrix(LngRow, MI_��ԭѧ���뼰����), True, "-") & "'," & zlStr.To_Date(.TextMatrix(LngRow, MI_�ͼ�����), "ymdhm") & ")"
                End If
            Next
        End With
    End If
End Sub

Private Function Get��ҳ����SQL(ByVal intType As Integer, ByVal arrFilds As Variant) As String
'���ܣ���ȡ��ҳ����Ĵ洢���̵���SQL
'������intType=0-������Ϣ��ҳ����SQL,1-������ҳ��ҳ����SQL
'      arrFilds=�ֶ���������
'���أ�SQL
    Dim strSql As String, strValue As String
    Dim i As Long, lngIdex As Long
    Dim lngMedType As Long
    
    If intType = 0 Then
        If Not gclsPros.OnlyPatiInfo And gclsPros.NoType = IT_New Or Not gclsPros.IsExistPati Then
            lngMedType = 1 '���Ӳ�����Ϣ
        End If
        strSql = "ZL" & IIf(gclsPros.FuncType = f������ҳ, 3, "") & "_������Ϣ_��ҳ����(" & IIf(gclsPros.FuncType = f������ҳ, lngMedType & ",", "")
    Else
        If gclsPros.OpenMode <> EM_�༭ And Not gclsPros.Is��Ŀ Then
            lngMedType = 1 '���Ӳ�����ҳ
        End If
        If Not gclsPros.OnlyPatiInfo And gclsPros.NoType = IT_New Or Not gclsPros.IsExistPati Then
            lngMedType = 1 '���Ӳ�����ҳ
        End If
        strSql = "ZL" & IIf(gclsPros.FuncType = f������ҳ, 3, "") & "_������ҳ_��ҳ����(" & IIf(gclsPros.FuncType = f������ҳ, lngMedType & ",", "")
    End If
    
    For i = LBound(arrFilds) To UBound(arrFilds)
        strValue = ""
        Select Case Trim(arrFilds(i))
            Case ""
                strValue = ",Null"
            Case "����Ա���", "����Ա����"
                strValue = ",'" & IIf(arrFilds(i) = "����Ա���", UserInfo.���, UserInfo.����) & "'"
            Case "����ID"
                strValue = gclsPros.����ID & ""
            Case "��ҳID", "��Ժ����ID", "��Ժ����ID"
                'arrFilds(i) & "",��֪��ʲôԭ��Decode���ֵ�һ������Ϊһ�����ֵ���Decode�������ƴ�ӿմ�
                strValue = "," & decode(arrFilds(i) & "", "��ҳID", gclsPros.��ҳID, "��Ժ����ID", ZVal(gclsPros.��Ժ����ID), "��Ժ����ID", gclsPros.��Ժ����ID)
            Case "NO"
                    strValue = IIf(gclsPros.RegistNo = "", ",Null", ",'" & gclsPros.RegistNo & "'")
            Case "��ͬ��λID"
                If Trim(gclsPros.CurrentForm.txtAdressInfo(ADRC_��λ��ַ).Text) <> "" Then
                    strValue = Val(gclsPros.CurrentForm.txtAdressInfo(ADRC_��λ��ַ).Tag)
                End If
                strValue = "," & ZVal(strValue)
            Case Else
                gclsPros.MainInfoRec.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                strValue = gclsPros.MainInfoRec!��Ϣ��ֵ & ""
                Select Case arrFilds(i)
                    Case "��������", "����ʱ��"
                        strValue = "," & zlStr.To_Date(strValue, "ymdhm")
                    Case "��Ժ����", "��Ժ����", "ȷ������"
                        strValue = "," & zlStr.To_Date(strValue, "ymdhms")
                    Case "��Ŀ����"
                        strValue = "," & zlStr.To_Date(strValue, "ymd")
                    Case "���", "����", "����Ժ"
                        strValue = "," & ZVal(strValue)
                    Case "����", "���ú�", "סԺ����", "�·�����", "�Ƿ�ȷ��"
                        strValue = "," & Val(strValue)
                    Case "ʬ���־"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).Text = "-", ",Null", "," & Val(strValue))
                    Case "�ɹ�����", "���ȴ���", "��������", "�����־"
                        strValue = IIf(strValue = "", ",Null", "," & Val(strValue))
                    Case "Ѫ��"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_Ѫ��).Text = "-", ",Null", ",'" & strValue & "'")
                    Case "RH"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_RH).Text = "-", ",Null", ",'" & strValue & "'")
                    Case Else
                        strValue = IIf(strValue = "", ",Null", ",'" & strValue & "'")
                End Select
        End Select
        If i = UBound(arrFilds) Then strValue = IIf(strValue = "", "Null", strValue) & ")"
        strSql = strSql & strValue
    Next
    Get��ҳ����SQL = strSql
    
End Function

Public Function CheckDateRange(ByVal strDate As String, Optional ByVal blnCheckData As Boolean) As Boolean
'���ܣ����¼�������Ƿ������Ժ���ڷ�Χ
'������strDate=����������
'      blnCheckData=true:ֻ������ڷ�Χ�������ʱ�䷶Χ��false:������ʱ�䷶Χ
'���أ�True=�ɹ��������Ժ�ڼ� �� false=ʧ�ܣ��������Ժ֮��
'˵������Ժ����Ϊ�գ�����false,��Ժ����Ϊ������Ϊ��ǰʱ��
    
    Dim DateStart As Date, dateEnd As Date
    Dim str��Ժʱ�� As String, str��Ժʱ�� As String
    Dim strFMT As String
    On Error GoTo errH
    
    CheckDateRange = False
    If Not IsDate(strDate) Then Exit Function
    Select Case Len(strDate)
        Case 10
            strFMT = "yyyy-MM-dd"
        Case 16
            strFMT = "yyyy-MM-dd hh:mm"
        Case 19
            strFMT = "yyyy-MM-dd hh:mm:ss"
        Case Else
            strFMT = "yyyy-MM-dd hh:mm"
    End Select
    '��ȡĬ�ϵ����Ժʱ��
    If Not IsDate(gclsPros.InTime) Then
        str��Ժʱ�� = "0"
    Else
        str��Ժʱ�� = Format(gclsPros.InTime, strFMT)
    End If
    If Not IsDate(gclsPros.OutTime) Then
        str��Ժʱ�� = "0"
    Else
        str��Ժʱ�� = Format(gclsPros.OutTime, strFMT)
    End If

    '��ʼʱ���ȡ
    DateStart = CDate(str��Ժʱ��)
    If DateStart = CDate(0) Then DateStart = zlDatabase.Currentdate
    '��ֹʱ���ȡ
    dateEnd = CDate(str��Ժʱ��)
    If dateEnd = CDate(0) Then dateEnd = zlDatabase.Currentdate
    
    'ʱ����
    If blnCheckData Then
        strDate = Format(strDate, "yyyy-MM-dd")
        If CDate(strDate) >= CDate(Format(DateStart, "yyyy-MM-dd")) And CDate(strDate) <= CDate(Format(dateEnd, "yyyy-MM-dd")) Then
            CheckDateRange = True
        End If
    Else
        If CDate(strDate) >= DateStart And CDate(strDate) <= dateEnd Then
            CheckDateRange = True
        End If
    End If
    
    Exit Function
errH:
    CheckDateRange = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetCboFromRec(ByVal arrIndex As Variant, Optional ByVal intInfoType As Integer, Optional ByVal strAddBeginItems As String = "NULL", Optional ByVal strAddEndItems As String = "NULL")
'���ܣ���ָ������Դ�е�����װ��ָ��������һ������ComboBox
'������arrIndex=ComboBox��Index���飬����ָ������Ϣ������ʱΪ�ַ�����Ϊ�˷�����չ�÷���)
'      intInfoType=0-�����ֵ����1-��Ա��
'      strDeFault=����б��д���Ĭ�ϱ�־����Ĭ��ֵ����Ч
'      strAddBeginItems=�Ƿ����б�ͷ�����µ�������Ŀ����";"�ָĬ��ֵNULL��ʶ�����
'      strAddEndItems=�Ƿ����б��β�����µ�������Ŀ����";"�ָĬ��ֵNULL��ʶ�����
    '����Ĵ���
    Dim arrItem As Variant
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
    Dim rsTmp As ADODB.Recordset
    '��ֱ�Ӵ�ֵ��ת��Ϊ����
    If TypeName(arrIndex) <> "Variant()" Then
        arrIndex = Array(arrIndex)
    End If
    If TypeName(arrIndex) = "Variant()" Then
        For i = LBound(arrIndex) To UBound(arrIndex)
            If intInfoType = 0 Then
                Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrIndex(i))
            Else
                Set objCboTmp = gclsPros.CurrentForm.cboManInfo(arrIndex(i))
            End If
            '��ӻ����¼��
            If intInfoType = 0 Then
                Set rsTmp = GetBaseCode(arrIndex(i))
            ElseIf intInfoType = 1 Then
                Set rsTmp = GetManData(arrIndex(i))
            End If
            '���ԭ������
            objCboTmp.Clear
            '��Ӷ����������
            If strAddBeginItems <> "NULL" Then
                arrItem = Split(strAddBeginItems, ",")
                For j = LBound(arrItem) To UBound(arrItem)
                    objCboTmp.AddItem arrItem(j)
                Next
            End If
            'װ������
            If Not rsTmp.EOF Then
                If objCboTmp.Index = BCC_Ѫ�� Then
                    objCboTmp.AddItem "-"
                End If
                For j = 1 To rsTmp.RecordCount
                    If IsNull(rsTmp!����) Then
                        objCboTmp.AddItem rsTmp!����
                    Else
                        objCboTmp.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����
                    End If
                    objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
                    If Val(rsTmp!ȱʡ & "") = 1 Then
                        Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                        objCboTmp.Tag = objCboTmp.NewIndex
                    End If
                    rsTmp.MoveNext
                Next
            End If
        Next
        '��Ӷ����������
        If strAddEndItems <> "NULL" Then
            arrItem = Split(strAddEndItems, ",")
            For j = LBound(arrItem) To UBound(arrItem)
                objCboTmp.AddItem arrItem(j)
                If intInfoType = 1 Then
                    objCboTmp.ItemData(objCboTmp.NewIndex) = -1
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub SetLstBoxFromRec(ByVal strlstInfos As String)
'���ܣ���ʼ��LVW
'������strlvwInfo�������Ϣ����,�����Ϣ����֮���Զ��ŷָ�
    Dim objlstBox As ListBox
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    Dim i As Long
    Dim blnDo As Boolean
    
    arrTmp = Split(strlstInfos, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Select Case arrTmp(i)
            Case "��Ⱦ��λ"
                Set objlstBox = gclsPros.CurrentForm.lstInfectParts
            Case "��Ⱦ����"
                Set objlstBox = gclsPros.CurrentForm.lstInfection
            Case "�����¼�"
                Set objlstBox = gclsPros.CurrentForm.lstAdvEvent
        End Select
        Set rsTmp = GetBaseCode(arrTmp(i))
        objlstBox.Clear
        rsTmp.Sort = "����,����"
        blnDo = arrTmp(i) = "�����¼�" And gclsPros.Is����
        Do While Not rsTmp.EOF
            If (rsTmp!���� & "" = "����������" Or rsTmp!���� & "" = "���������������") Then
               If blnDo Then
                   objlstBox.AddItem rsTmp!����
                   objlstBox.ItemData(objlstBox.NewIndex) = Val(rsTmp!����)
               End If
            Else
               objlstBox.AddItem rsTmp!����
               objlstBox.ItemData(objlstBox.NewIndex) = Val(rsTmp!����)
            End If
            rsTmp.MoveNext
        Loop
    Next
    objlstBox.ListIndex = -1
End Sub

Public Function SetInputRoot(ByVal intType As Integer, ByVal intSysPara As Integer, ByRef intModPara As Integer, ParamArray arrControls() As Variant) As Boolean
'˵�����ú�������ϵͳ������ģ�������ͬ����һ�鵥ѡ��ť��ϵͳ����ֵһ��ΪA(0��1),A+1,A+2....,ģ�����ΪB,B+1,....ϵͳ����ΪAʱ��ģ�����������,������������
'           ģ�����=B(ϵͳ����=A)������ҵ��Ч����ϵͳ����=A+1��ͬ
'           ģ�����=B+1(ϵͳ����=A)������ҵ��Ч����ϵͳ����=A+2��ͬ
'���ܣ�������Դ����������ģ�������ҽ�����Դ����ҽ�����Դ������������Դ
'������intType=0-��ҽ�����Դ���ã�1-��ҽ�����Դ��2-���������Դ
'      intSysPara=ϵͳ����������ֵΪA(0��1),A+1,A+2��..��ֵΪAʱģ�����������
'      intModPara=ģ�����
'���أ��Ƿ�ɹ�
'      intModPara=ʵ�ʲ���ֵ����ϵͳ����Ϊ��0��1��2��ģ��Ϊ0��1 ��ϵͳΪ0ʱģ�������ã���ʱģ�����ʵ��ֵ=ģ�����ֵ����ϵͳ����<>0����1��ģ�����ʵ��ֵ=ϵͳ����-1

    Dim blnVisual As Boolean, blnEnable As Boolean
    Dim i As Long
    Dim blnAller As Boolean

    On Error GoTo errH
    '����������Դ����������̫Ԫͨʱ�ؼ����ɼ�,��������ɼ�
    blnVisual = intType = 2 And gclsPros.PassType = 3 Or intType <> 2
    blnEnable = intSysPara = IIf(intType <> 2, 1, 0)
    If Not blnVisual Then intModPara = 0
    If Not blnEnable Then intModPara = intSysPara - IIf(intType <> 2, 2, 1)
    '���ÿؼ���ֵ�Լ�������
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Visible = blnVisual
        If blnVisual Then
            arrControls(i).Enabled = blnEnable And arrControls(i).Enabled
            'ʵ��ģ�����ֵ��ؼ������±���ʼֵһ����˳��һ��
            If i = intModPara Then
                arrControls(i).Value = 1
            Else
                arrControls(i).Value = 0
            End If
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SetCtrlValues(ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal str���ӱ��� As String, Optional ByVal blnMain As Boolean) As Boolean
'���ܣ����ÿؼ�ֵ
'����  strInfoName=��Ϣ��
'      strInfoValue=��Ϣֵ
'      str���ӱ���=����������Ŀ�����ж�
    Dim str�ؼ��� As String, strFMT As String
    Dim lngCount As Long, i As Long, j As Long, LngRow As Long
    Dim arrTmp As Variant, strTmp As String
    Dim vsTmp As VSFlexGrid, lstTmp As ListBox
    Dim intIndex As Integer, intIndexTmp As Integer
    Dim LngCols As Long

    On Error GoTo errH
    '����95480
    '������Ŀ�������������ԭ�����ڵĲ�����ҳ�ӱ���Ϣ�������Ƴ�ͻ��
    '������ȼ��ز���������Ŀ��Ȼ���ٵ���������Ƿ��и�����Ϣ
    If str���ӱ��� <> "" And gclsPros.FuncType <> f���ѡ�� Then
        Set vsTmp = gclsPros.CurrentForm.vsfMain
        LngCols = 6
        With vsTmp
            For i = 0 To LngCols Step 3
                LngRow = -1: LngRow = .FindRow(strInfoName, , i)
                If LngRow >= 0 Then
                    If .TextMatrix(LngRow, i + 2) = "�Ƿ�" Then
                        .Cell(flexcpChecked, LngRow, i + 1) = IIf(Val(strInfoValue) = 0, 2, 1)
                    Else
                        .TextMatrix(LngRow, i + 1) = strInfoValue
                    End If
                    Call UpdateCacheRecInfo(0, "������Ŀ", strInfoValue, strInfoValue, LngRow, , .TextMatrix(LngRow, i) & ";" & LngRow & ";" & i)
                    Exit For
                End If
            Next
        End With
    ElseIf str���ӱ��� = "" Then
        Select Case strInfoName
            Case "����ʱ��", "�������", "̥��", "̥��", "����ʱ��1", "����ʱ��2", "����ʱ��3", _
                    "�ܲ���ʱ��", "�����Ѫ��", "���Ʋ���֢", "�����������"
                If grsDeliceryInfo Is Nothing Then Exit Function  'ֻ��Ҫ����
                grsDeliceryInfo.AddNew Array("��Ϣ��", "��Ϣֵ", "��Ϣ��ֵ", "����"), Array(strInfoName, strInfoValue, strInfoValue, 0)
                grsDeliceryInfo.Update
                Exit Function  'ֻ��Ҫ����
            Case "������4", "������5", "������6"
                gclsPros.MainInfoRec.Filter = "��Ϣ��='������'"
            Case "CT", "PETCT", "˫ԴCT", "XƬ", "B��", "�����Ķ�ͼ", "MRI", "ͬλ�ؼ��"
                If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                    gclsPros.MainInfoRec.Filter = "��Ϣ��='������'"
                Else
                    gclsPros.MainInfoRec.Filter = "��Ϣ��='" & strInfoName & "'"
                End If
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
            End If
        Else
            If strInfoName = "סԺ����" And gclsPros.FuncType = f������ҳ And gclsPros.OpenMode <> EM_�������� Then Exit Function
            str�ؼ��� = gclsPros.MainInfoRec!�ؼ��� & ""
            With gclsPros.CurrentForm
                '������Ϣ��չ״̬
                If gclsPros.MainInfoRec!ExpState = 0 Then
                    intIndex = Val(gclsPros.MainInfoRec!Index & "")
                    Select Case str�ؼ���
                        Case "txtSpecificInfo"
                            Select Case intIndex
                                Case SLC_Ӥ�׶�����, SLC_����
                                    Call LoadOldData(strInfoValue, intIndex)
                                Case SLC_��֢�໤��, SLC_��֢�໤Сʱ
                                    .txtSpecificInfo(intIndex).Text = strInfoValue
                                    .optInput(OP_ICU��).Value = 1
                                Case SLC_Ժ�ڻ���, SLC_��Ժ����
                                    .txtSpecificInfo(intIndex).Text = Val(strInfoValue)
                                    .chkInfo(CHK_�������).Value = 1
                                Case Else
                                    .txtSpecificInfo(intIndex).Text = strInfoValue
                            End Select
                        Case "cboBaseInfo"
                            '����������
                            If blnMain And strInfoValue = "" Then '������ҳ������Ϣ����ϢΪ�գ�������Ĭ��ֵ��Ĭ��ֵ�б����ʱ�Ѿ����ã�
                                If strInfoName = "ʬ���־" Then
                                    strInfoValue = "-"
                                    If gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).ListCount >= 1 Then
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).Clear
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "-"
                                    Else
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).Clear
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "0-��"
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "1-��"
                                    End If
                                    Call Cbo.SeekIndex(.cboBaseInfo(intIndex), strInfoValue)
                                End If
                            Else
                                If strInfoName = "Ѫ��" Then
                                    If strInfoValue = "" Then
                                        strInfoValue = "-"
                                    ElseIf strInfoValue = "δ֪" Then
                                        strInfoValue = "����" 'δ֪ ��Ϊ ��
                                    Else
                                        strInfoValue = strInfoValue
                                    End If
                                ElseIf strInfoName = "RH" Then
                                    If strInfoValue = "" Then
                                        strInfoValue = "-"
                                    ElseIf strInfoValue = "δ��" Then
                                        strInfoValue = "δ��" 'δ�� ��Ϊ δ��
                                    Else
                                        strInfoValue = strInfoValue
                                    End If
                                End If
                                If strInfoName = "����Ժ�ƻ�����" Or strInfoName = "ʬ���־" Then
                                    .cboBaseInfo(intIndex).ListIndex = Val(strInfoValue)
                                    If Val(strInfoValue) = 0 Then strInfoValue = ""
                                Else
                                    '����Index��ؼ�ֵ
                                    Call Cbo.SeekIndex(.cboBaseInfo(intIndex), strInfoValue)
                                    If .cboBaseInfo(intIndex).ListIndex = -1 And strInfoValue <> "" Then
                                        If .cboBaseInfo(intIndex).Style = 0 Then
                                            .cboBaseInfo(intIndex).Text = strInfoValue
                                        Else
                                            '����ϵͳ��ǰ���ܶ����в��淶��ֵ
    '                                        If strInfoName = "��������" Or strInfoName = "ȥ��" Then
                                            Call SetCboFromName(strInfoValue, .cboBaseInfo(intIndex), , True)
                                        End If
                                    End If
                                End If
                            End If
                            If intIndex = BCC_���֤ Then '���֤�ؼ��洢������Ϣ��
                                If zlCommFun.ActualLen(strInfoValue) = Len(strInfoValue) Then
                                    If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                                        strInfoValue = IIf(strInfoName = "���֤��״̬", "", strInfoValue)
                                        If zlStr.ActualLen(strInfoValue) > 12 And gclsPros.IsMaskID Then   '�������֤������
                                            .cboBaseInfo(intIndex).Tag = "������Change�¼�" '��ǲ�����Change�¼�
                                            .cboBaseInfo(intIndex).Text = Mid(strInfoValue, 1, 12) & String(Len(Mid(strInfoValue, 13, 2)), "*") & Mid(strInfoValue, 15)
                                            .cboBaseInfo(intIndex).Tag = strInfoValue
                                        End If
                                    Else
                                         strInfoValue = IIf(strInfoName = "�⼮���֤��", strInfoValue, "")
                                    End If
                                Else '�������ģ���Ϊ���֤��״̬
                                    strInfoValue = IIf(strInfoName = "���֤��״̬", zlCommFun.GetNeedName(strInfoValue), "")
                                End If
                            End If
    
                        Case "txtInfo"
                            .txtInfo(intIndex).Text = strInfoValue
                            intIndexTmp = decode(strInfoName, "�˳�ԭ��", CHK_���·��, "����ԭ��", CHK_����, "�������", CHK_�������, -1)
                            If intIndexTmp <> -1 Then
                                If strInfoValue = IIf(strInfoName <> "�������", "1", "0") Then
                                    .chkInfo(intIndexTmp).Value = Val(strInfoValue)
                                    .txtInfo(intIndex).Text = ""
                                Else
                                    If strInfoValue <> "" And strInfoName = "����ԭ��" Then
                                        .chkInfo(intIndexTmp).Value = 1
                                        If gclsPros.PathVCauses Then
                                            Call Cbo.SeekIndex(.cboBaseInfo(BCC_����ԭ��), strInfoValue)
                                            If .cboBaseInfo(BCC_����ԭ��).ListIndex = -1 Then
                                                .cboBaseInfo(BCC_����ԭ��).AddItem strInfoValue
                                                .cboBaseInfo(BCC_����ԭ��).ListIndex = .cboBaseInfo(BCC_����ԭ��).NewIndex
                                            End If
                                        End If
                                    Else
                                        .chkInfo(intIndexTmp).Value = IIf(strInfoName = "�˳�ԭ��", 0, 1)
                                    End If
                                End If
                            End If
                        Case "chkInfo"
                            .chkInfo(intIndex).Value = IIf(Val(strInfoValue) = 0, 0, 1)
                            Call chkInfoClick(intIndex)
                        Case "cboManInfo"
                            If strInfoName = "��ĿԱ����" And (strInfoValue = "" Or gclsPros.OpenMode = EM_�������� Or gclsPros.OpenMode = EM_������ҳ) Then
                                .cboManInfo(intIndex).Text = UserInfo.����
                            Else
                                .cboManInfo(intIndex).Text = strInfoValue
                            End If
                        Case "mskDateInfo"
                            'ʱ�䰴�ؼ���MASKֵ��ʼ��
                            strFMT = .mskDateInfo(intIndex).Mask
                            If gclsPros.FuncType = fҽ����ҳ And intIndex = DC_�������� Then
                                If Format(strInfoValue, "HH:MM") = "00:00" Then
                                    .mskDateInfo(intIndex).Mask = "####-##-##"
                                    .mskDateInfo(intIndex).Tag = "####-##-##"
                                    strFMT = .mskDateInfo(intIndex).Mask
                                 Else
                                    .mskDateInfo(intIndex).Mask = "####-##-## ##:##"
                                    .mskDateInfo(intIndex).Tag = "####-##-## ##:##"
                                    strFMT = .mskDateInfo(intIndex).Mask
                                End If
                            End If
                            If IsDate(strInfoValue) Then
                                strInfoValue = Format(strInfoValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            Else
                                strInfoValue = Replace(strFMT, "#", "_")
                            End If
                            .mskDateInfo(intIndex).Text = strInfoValue
                            If Not IsDate(strInfoValue) Then strInfoValue = ""
                            If strInfoValue = "" And intIndex = DC_��Ŀ���� Then
                                .mskDateInfo(intIndex).Text = Format(zlDatabase.Currentdate, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            End If
                            .txtDateInfo(intIndex).Text = .mskDateInfo(intIndex).Text
                            If intIndex = DC_�������� And gclsPros.FuncType = f������ҳ Then
                                If Trim(.mskDateInfo(intIndex).Text) = "____-__-__ __:__" Then
                                    Call SetCtrlLocked(.mskDateInfo(intIndex), True)
                                    Call SetCtrlLocked(.txtDateInfo(intIndex), True)
                                End If
                            End If
                        Case "txtAdressInfo"
                            Call SetPatiAddress(intIndex, strInfoName, strInfoValue)
                        Case "cboSpecificInfo"
                            Call Cbo.SeekIndex(.cboSpecificInfo(intIndex), strInfoValue)
                            If .cboSpecificInfo(intIndex).Style = 0 And .cboSpecificInfo(intIndex).ListIndex = -1 Then
                               .cboSpecificInfo(intIndex).Text = strInfoValue
                            End If
                        Case "lstInfection", "lstAdvEvent", "lstInfectParts"
                            If strInfoName = "��Ⱦ����" Then
                                Set lstTmp = .lstInfection
                            ElseIf strInfoName = "�����¼�" Then
                                Set lstTmp = .lstAdvEvent
                            ElseIf strInfoName = "��Ⱦ��λ" Then
                                Set lstTmp = .lstInfectParts
                            End If
                            If InStr(strInfoValue, ",") > 0 Then
                                strInfoValue = Replace(strInfoValue, ",", "|") '�����ŷָ����ת��Ϊ��|��
                            End If
                            arrTmp = Split(strInfoValue, "|")
                            For j = 0 To lstTmp.ListCount - 1
                                For i = LBound(arrTmp) To UBound(arrTmp)
                                    If lstTmp.ItemData(j) = arrTmp(i) Then
                                        lstTmp.Selected(j) = True: Exit For
                                    End If
                                Next
                            Next
                            lstTmp.ListIndex = -1
                    End Select
                    gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                ElseIf gclsPros.MainInfoRec!ExpState = ES_��ʼ��չ Then
                    If str�ؼ��� <> "vsTSJC" Then
                        gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                        gclsPros.SecdInfoRec.Sort = "Sort"
                    End If
                    Select Case strInfoName
                        Case "����ʱ��"
                            '�����ʽ:��Ժǰ(�죬Сʱ,����)|��Ժ��(�죬Сʱ,����)
                            strTmp = Replace(strInfoValue, "|", ",")
                            strTmp = strTmp & ",,,,,"
                            arrTmp = Split(strTmp, ",")
                            For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                .txtSpecificInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                                gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(arrTmp(i), arrTmp(i))
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                        Case "ת�Ƽ�¼"
                            strTmp = strInfoValue & ",,,,,,"
                            arrTmp = Split(strTmp, ",")
                            If str�ؼ��� = "txtInfo" Then
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                                    gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            Else
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .vsTransfer.TextMatrix(DR_ת�ƿ���, Val(gclsPros.SecdInfoRec!IndexEx & "")) = arrTmp(i)
                                    gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            End If
                            gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                        Case "ת��ʱ��"
                            strTmp = strInfoValue & ",,,,,,"
                            arrTmp = Split(strTmp, ",")
                            If str�ؼ��� = "txtInfo" Then
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = zlStr.FullDate(arrTmp(i), True)
                                    gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            Else
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .vsTransfer.TextMatrix(DR_ת��ʱ��, Val(gclsPros.SecdInfoRec!IndexEx & "")) = zlStr.FullDate(arrTmp(i), True)
                                    gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            End If
                            gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                        Case "����ʱ��"
                            For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                strFMT = .mskDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Mask
                                If IsDate(strInfoValue) Then
                                    strTmp = Format(strInfoValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "##:##", "HH:mm"))
                                    If strTmp = "00:00" Then strTmp = Replace(strFMT, "#", "_")
                                Else
                                    strTmp = Replace(strFMT, "#", "_")
                                End If
                                .mskDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = strTmp
                                .txtDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = strTmp
                                If Not IsDate(strTmp) Then strTmp = ""
                                gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                        Case "31������סԺ"
                            .optInput(OP_��סԺ��).Value = strInfoValue = ""
                            .optInput(OP_��סԺ��).Value = strInfoValue <> ""
                            .txtInfo(GC_31������סԺ).Text = strInfoValue
                            Call SetCtrlLocked(.txtInfo(GC_31������סԺ), strInfoValue = "", True)
                            For i = 1 To gclsPros.SecdInfoRec.RecordCount
                                strTmp = decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), OP_��סԺ��, IIf(strInfoValue = "", 1, 0), OP_��סԺ��, IIf(strInfoValue = "", 0, 1), strInfoValue)
                                gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                        Case "����"
                            .optState(OP_����).Value = Val(strInfoValue) = 0
                            .optState(OP_����).Value = Val(strInfoValue) <> 0
                            For i = 1 To gclsPros.SecdInfoRec.RecordCount
                                strTmp = IIf(Val(gclsPros.SecdInfoRec!IndexEx & "") = OP_����, IIf(Val(strInfoValue) <> 0, 1, 0), IIf(Val(strInfoValue) <> 0, 0, 1))
                                gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                        Case Else
                            If str�ؼ��� = "vsTSJC" Then
                                If strInfoName Like "������*" And gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                                    intIndex = Val(Mid(strInfoName, 5, 1)) - 4
                                    strTmp = strInfoValue
                                Else
                                    intIndex = decode(strInfoName, "CT", TR_CT, "PETCT", TR_PETCT, "˫ԴCT", TR_˫ԴCT, _
                                                "XƬ", TR_XƬ, "B��", TR_B��, "�����Ķ�ͼ", TR_�����Ķ�ͼ, "MRI", TR_MRI, "ͬλ�ؼ��", TR_ͬλ�ؼ��, -1)
                                    strTmp = decode(Val(strInfoValue), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
                                End If
                                If intIndex <> -1 Then
                                strInfoValue = strTmp
                                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And IndexEx=" & intIndex
                                    .vsTSJC.TextMatrix(intIndex, 1) = strTmp
                                    .vsTSJC.Cell(flexcpData, intIndex, 1) = strTmp
                                    gclsPros.SecdInfoRec.Update Array("��Ϣԭֵ", "����Ϣԭֵ"), Array(strInfoValue, strInfoValue)
                                End If
                            End If
                    End Select
                    If str�ؼ��� <> "vsTSJC" Then
                        gclsPros.MainInfoRec.Update "��Ϣԭֵ", strInfoValue
                    End If
                ElseIf gclsPros.MainInfoRec!ExpState = 2 Then
                '����ʱ����
                End If
            End With
        End If
    End If
    SetCtrlValues = True
    Exit Function
errH:
    Debug.Print "SetCtrlValues:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CacheCtrlValues()
'���ܣ��������ؼ�ֵ���ؼ���Ӧ��Ϣһ�㲻��չ��μ���չ
    Dim str�ؼ��� As String, strInfoValue As String, strInfoName As String
    Dim blnEnable As Boolean
    Dim lstTmp As ListBox
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim vsTmp As VSFlexGrid
    Dim intIndex As Integer
    Dim LngCols As Long

    On Error GoTo errH
    With gclsPros.CurrentForm
        '����չ��Ϣ�Ѽ�����
        gclsPros.MainInfoRec.Filter = "ExpState=0": gclsPros.MainInfoRec.Sort = "���"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            strInfoValue = ""
            str�ؼ��� = gclsPros.MainInfoRec!�ؼ���
            strInfoName = gclsPros.MainInfoRec!��Ϣ��
            Select Case str�ؼ���
                Case "txtInfo"
                    Select Case strInfoName
                        Case "�˳�ԭ��", "����ԭ��"
                            If .chkInfo(CHK_����·��).Value = 1 Then
                                intIndex = IIf(strInfoName = "�˳�ԭ��", CHK_���·��, CHK_����)
                                If .chkInfo(intIndex).Value = 1 Then
                                    strInfoValue = IIf(strInfoName = "�˳�ԭ��", "1", .txtInfo(gclsPros.MainInfoRec!Index).Text)
                                    'û����д����ԭ���򱣴�1
                                    If strInfoValue = "" And strInfoName = "����ԭ��" Then strInfoValue = "1"
                                Else
                                    strInfoValue = IIf(strInfoName = "�˳�ԭ��", .txtInfo(gclsPros.MainInfoRec!Index).Text, "")
                                End If
                            End If
                        Case "�������"
                            If .chkInfo(CHK_�������).Value = 1 Then
                                strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                            Else
                                strInfoValue = "0"
                            End If
                        Case "��ϵ�˸�����Ϣ"
                            If .txtInfo(gclsPros.MainInfoRec!Index).Visible Then
                                strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                            End If
                        Case Else
                            strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                    End Select
                Case "txtSpecificInfo"
                    strInfoValue = .txtSpecificInfo(gclsPros.MainInfoRec!Index).Text
                    Select Case strInfoName
                        Case "����"
                            '��֪Ϊʲô�û�����������ҽ���ҳ�ж�.cboSpecificInfo(gclsPros.MainInfoRec!Index).VisibleΪFalse
                            If strInfoValue <> "" Then strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text, "")
                        Case "������������"
                            If strInfoValue <> "" Then
                                If .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text = "��" Then
                                    strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), "��" & Trim(.txtSpecificInfo(SLC_Ӥ�׶�����_DAY)) & "��", "")
                                Else
                                    strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text, "")
                                End If
                            End If
                        Case "��������"
                            If .chkInfo(CHK_����).Value = 1 Then
                                strInfoValue = IIf(Val(strInfoValue) <> 0, Val(strInfoValue), "")
                            Else
                                strInfoValue = ""
                            End If
                        Case "Ժ�ڻ���", "��Ժ����"
                            If .chkInfo(CHK_�������).Value = 0 Then
                                strInfoValue = ""
                            End If
                    End Select
                Case "chkInfo"
                    strInfoValue = .chkInfo(gclsPros.MainInfoRec!Index).Value
                    If strInfoValue = "0" Then strInfoValue = ""
                    If strInfoName = "�����־" Then
                        If strInfoValue = "1" Then
                            strInfoValue = decode(zlStr.NeedName(.cboSpecificInfo(SLC_��������).Text), "��", 1, "��", 2, "��", 3, "��", 4, "����", 9)
                        End If
                    End If
                Case "cboBaseInfo"
                    strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).Text
                    If strInfoName = "��������" Or strInfoName = "����״��" Then
                        If InStr(strInfoValue, "-") > 0 Then  '����ǹ淶�����ݣ���ֻ�����
                            strInfoValue = Mid(strInfoValue, 1, InStr(strInfoValue, "-") - 1)
                        End If
                    ElseIf strInfoName = "��Ѫ��Ӧ" Then
                        strInfoValue = IIf(.cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex = -1, "", .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex)
                    ElseIf strInfoName = "����Ժ�ƻ�����" Or strInfoName = "ʬ���־" Then
                        If .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex > 0 Then
                            strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex
                        Else
                            strInfoValue = ""
                        End If
                    ElseIf Val(gclsPros.MainInfoRec!Index & "") = BCC_���֤ Then '���֤�ؼ��洢������Ϣ��
                        If zlCommFun.ActualLen(strInfoValue) = Len(strInfoValue) Then
                            If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                                If gclsPros.FuncType = fҽ����ҳ And strInfoValue <> "" Then '��������Ĵ��ڣ�����ȡTag
                                    strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).Tag
                                End If
                                strInfoValue = IIf(strInfoName = "���֤��", strInfoValue, "")
                            Else
                                strInfoValue = IIf(strInfoName = "�⼮���֤��", strInfoValue, "")
                            End If
                        Else '�������ģ���Ϊ���֤��״̬
                            strInfoValue = IIf(strInfoName = "���֤��״̬", zlCommFun.GetNeedName(strInfoValue), "")
                        End If
                    Else
                        strInfoValue = zlStr.NeedName(strInfoValue)
                    End If
                Case "cboManInfo"
                    strInfoValue = zlStr.NeedName(.cboManInfo(gclsPros.MainInfoRec!Index).Text)
                Case "txtAdressInfo"
                    On Error Resume Next
                    strInfoValue = .padrInfo(gclsPros.MainInfoRec!Index).Value
                    If Err.Number <> 0 Then
                        Err.Clear: strInfoValue = .txtAdressInfo(gclsPros.MainInfoRec!Index).Text
                    Else
                        If Not .padrInfo(gclsPros.MainInfoRec!Index).Visible Then
                            strInfoValue = .txtAdressInfo(gclsPros.MainInfoRec!Index).Text
                        End If
                    End If
                    On Error GoTo errH
                Case "mskDateInfo"
                    strInfoValue = .mskDateInfo(gclsPros.MainInfoRec!Index).Text
                    If Not IsDate(strInfoValue) Then strInfoValue = ""
                Case "cboSpecificInfo"
                    strInfoValue = .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text
                Case "lstInfection", "lstAdvEvent", "lstInfectParts"
                    If strInfoName = "��Ⱦ����" Then
                        Set lstTmp = .lstInfection
                    ElseIf strInfoName = "�����¼�" Then
                        Set lstTmp = .lstAdvEvent
                    ElseIf strInfoName = "��Ⱦ��λ" Then
                        Set lstTmp = .lstInfectParts
                    End If
                    For j = 0 To lstTmp.ListCount - 1
                        If lstTmp.Selected(j) = True Then
                            strInfoValue = strInfoValue & "|" & lstTmp.ItemData(j)
                        End If
                    Next
                    If strInfoValue <> "" Then
                        strInfoValue = Mid(strInfoValue, 2)
                    End If
            End Select
            gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
            gclsPros.MainInfoRec.MoveNext
        Next
        '�μ���չ��Ϣ�Ѽ�����
        gclsPros.MainInfoRec.Filter = "ExpState=1": gclsPros.MainInfoRec.Sort = "���"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            str�ؼ��� = gclsPros.MainInfoRec!�ؼ��� & ""
            strInfoName = gclsPros.MainInfoRec!��Ϣ��
            strInfoValue = ""
            Select Case strInfoName
                Case "����ʱ��"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        strTmp = .txtSpecificInfo(gclsPros.SecdInfoRec!IndexEx).Text
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & IIf(j = 4, "|", ",") & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
                Case "ת�Ƽ�¼"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        If str�ؼ��� = "txtInfo" Then
                            strTmp = .txtInfo(gclsPros.SecdInfoRec!IndexEx).Text
                        Else
                            strTmp = .vsTransfer.TextMatrix(DR_ת�ƿ���, gclsPros.SecdInfoRec!IndexEx)
                        End If
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & "," & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    If strInfoValue <> "" Then strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
                Case "ת��ʱ��"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        If str�ؼ��� = "txtInfo" Then
                            strTmp = Format(.txtInfo(gclsPros.SecdInfoRec!IndexEx).Text, "yyyyMMddHHmm")
                        Else
                            strTmp = Format(.vsTransfer.TextMatrix(DR_ת��ʱ��, gclsPros.SecdInfoRec!IndexEx), "yyyyMMddHHmm")
                        End If
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & "," & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    If strInfoValue <> "" Then strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
                Case "����ʱ��"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                        If j = 0 Then
                            strInfoValue = .mskDateInfo(gclsPros.SecdInfoRec!IndexEx).Text
                            If Not IsDate(strInfoValue) Then strInfoValue = ""
                            Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strInfoValue, strInfoValue))
                        Else
                            If strInfoValue <> "" Then
                                strTmp = .mskDateInfo(gclsPros.SecdInfoRec!IndexEx).Text
                                If IsDate(strTmp) Then
                                    Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strTmp, strTmp))
                                    strInfoValue = strInfoValue & " " & strTmp
                                End If
                            End If
                        End If
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
                Case "31������סԺ"
                        gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                        gclsPros.SecdInfoRec.Sort = "Sort"
                        strInfoValue = .txtInfo(GC_31������סԺ).Text
                        If .optInput(OP_��סԺ��).Value = 1 Then strInfoValue = ""
                        For j = 1 To gclsPros.SecdInfoRec.RecordCount
                            strTmp = decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), OP_��סԺ��, IIf(strInfoValue = "", 1, 0), OP_��סԺ��, IIf(strInfoValue = "", 0, 1), strInfoValue)
                            Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strTmp, strTmp))
                            gclsPros.SecdInfoRec.MoveNext
                        Next
                        gclsPros.MainInfoRec.Update "��Ϣ��ֵ", strInfoValue
                Case "����"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        strInfoValue = IIf(Val(gclsPros.SecdInfoRec!IndexEx & "") = OP_����, IIf(.optState(OP_����).Value, 1, 0), IIf(.optState(OP_����).Value, 1, 0))
                        gclsPros.SecdInfoRec.Update Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strInfoValue, strInfoValue)
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    gclsPros.MainInfoRec.Update "��Ϣ��ֵ", IIf(.optState(OP_����).Value, 1, 0)
                Case "������"
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!���
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 0 To .vsTSJC.Rows - 1
                        strInfoValue = .vsTSJC.TextMatrix(gclsPros.SecdInfoRec!IndexEx, 1)
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(strInfoValue, strInfoValue))
                        gclsPros.SecdInfoRec.MoveNext
                    Next
            End Select
            gclsPros.MainInfoRec.MoveNext
        Next
    End With

    '��չ��Ϣ�Ѽ�
    gclsPros.MainInfoRec.Filter = "��Ϣ��='������Ŀ'"
    If Not gclsPros.MainInfoRec.EOF Then
        Set vsTmp = gclsPros.CurrentForm.vsfMain
        LngCols = 6
        With vsTmp
            For i = .FixedRows To .Rows - 1
                For j = 0 To LngCols Step 3
                    If .TextMatrix(i, j) <> "" Then
                        If .TextMatrix(i, j + 2) = "�Ƿ�" Then
                            strInfoValue = IIf(.Cell(flexcpChecked, i, j + 1) = 2, "", 1)
                        Else
                            strInfoValue = .TextMatrix(i, j + 1)
                        End If
                        Call UpdateCacheRecInfo(1, "������Ŀ", strInfoValue, strInfoValue, i, , .TextMatrix(i, j) & ";" & i & ";" & j)
                    End If
                Next
            Next
        End With
    End If
    Exit Sub
errH:
    Debug.Print "CacheCtrlValues:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub LoadOldData(ByVal strOld As String, Optional ByVal intIndex As Integer)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
'������strOld =�����ַ���
'      intIndex=��������Ŀؼ�����ֵֵ
    Dim strTmp As String, lngIdx As Long
    Dim objTxt As TextBox, objCbo As ComboBox
    Dim arrTmp As Variant
    Dim i As Long

    If Trim(strOld) = "" Then Exit Sub
    If intIndex = SLC_���� Then
        strTmp = "��,��,��,Сʱ,����"
    ElseIf intIndex = SLC_Ӥ�׶����� Then
        strTmp = "��,��,Сʱ,����"
    End If
    arrTmp = Split(strTmp, ",")

    strTmp = strOld
    For i = LBound(arrTmp) To UBound(arrTmp)
        If intIndex = SLC_Ӥ�׶����� And strTmp Like "*��*��" Then
            strTmp = Replace(strTmp, "��", "")
            gclsPros.CurrentForm.txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Text = Split(strTmp, "��")(1)
            strTmp = Split(strTmp, "��")(0)
            lngIdx = i
            Exit For
        ElseIf InStr(strOld, arrTmp(i)) > 0 Then
            If InStr(strOld, arrTmp(i)) + Len(arrTmp(i)) - 1 = Len(strOld) Then
                strTmp = Mid(strOld, 1, InStr(strOld, arrTmp(i)) - 1)
                lngIdx = i
            End If
            Exit For
        End If
    Next

    'IsNumeric("")=False,����������ַ���1
    If Not IsNumeric(strTmp & "1") Then
        lngIdx = -1
        strTmp = strOld
    End If
    
    '��֪Ϊʲô�û�����������ҽ���ҳ�ж�objCbo.VisibleΪFalse
    Set objTxt = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    Set objCbo = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
    objTxt.Text = strTmp

    If lngIdx = -1 Then
        objCbo.Visible = False
        objCbo.Tag = "����"
        objCbo.ListIndex = -1
        If objCbo.Container.Name = "fraCbo" Then
            objCbo.Container.Visible = False
        End If
        
        If intIndex = SLC_���� Then
            If gclsPros.FuncType = f������ҳ Then
                objTxt.Width = 1250
            Else
                objTxt.Width = 1150
            End If
        ElseIf intIndex = SLC_Ӥ�׶����� Then
            objTxt.Width = 1250
        End If
    Else
        If objCbo.Visible = False Then
            objCbo.Visible = True
            objCbo.Tag = ""
            If objCbo.Container.Name = "fraCbo" Then
                objCbo.Container.Visible = True
            End If
            
            If intIndex = SLC_���� Then
                If gclsPros.FuncType = f������ҳ Then
                    objTxt.Width = 450
                Else
                    objTxt.Width = 360
                End If
            ElseIf intIndex = SLC_Ӥ�׶����� Then
                objTxt.Width = 360
            End If
        End If
        objCbo.ListIndex = lngIdx
    End If
End Sub

Public Sub SetPageVisible()
'���ܣ�����ҳ��ɼ���
    Dim i As Long
    With gclsPros.CurrentForm
        '����ҽ�Ʋ���ʾ��ҽ���
        If gclsPros.PatiType = PF_סԺ Then
            For i = .PicPage.LBound To .PicPage.UBound
                .PicPage(i).Tag = "true"
            Next
             'û�б༭��ҽȨ�޵Ĳ���ʾ��ҽ
            .PicPage(PIC_��ҽ���).Tag = IIf(gclsPros.Have��ҽ, "true", "false")
            .PicPage(PIC_��ҽ������).Tag = IIf(gclsPros.Have��ҽ, "true", "false")
            Select Case gclsPros.MedPageSandard
                Case ST_����ʡ��׼
                    .PicPage(PIC_������).Tag = "false"
                    .PicPage(PIC_��֢�໤).Tag = "false"
                Case ST_����ʡ��׼
                    .PicPage(PIC_������).Tag = "false"
                    .PicPage(PIC_��֢�໤).Tag = "false"
                Case ST_�Ĵ�ʡ��׼
                    .PicPage(PIC_������).Tag = "false"
            End Select
            If gclsPros.FuncType = fҽ����ҳ Then
                .PicPage(PIC_סԺ����).Tag = "false"
            End If
            If Not gclsPros.ReadPages Then  '����������װʱ������ʾ�����뻯���Լ�����ҩƷ
                .PicPage(PIC_������Ϣ).Tag = "false"
                .PicPage(PIC_���Ƽ�¼).Tag = "false"
                If gclsPros.MedPageSandard = ST_��������׼ Then
                    .PicPage(PIC_������).Tag = "false"
                End If
            End If
            For i = .PicPage.LBound To .PicPage.UBound
                If .PicPage(i).Tag = "true" Then
                    .PicPage(i).Visible = True
                ElseIf .PicPage(i).Tag = "false" Then
                    .PicPage(i).Visible = False
                End If
            Next
            '���õ���Ŀ¼
            Call SetMainDirectory
        End If
    End With
End Sub

Public Function SetSignature() As Boolean
'���ܣ����ݵ�ǰ���˵�ҽʦ��ǩ�������ȷ��ǩ�����������ݵĿɱ༭��
'���أ������Ƿ���ǩ��ֻ�����ܱ༭
    Static rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim strSql As String, blnReadOnly As Boolean
    Dim i As Integer, j As Integer
    Dim strTmp As String
    '˵����arrInfos��arrManIdxs��arrSgnIdxs���������Ԫ��һһ��Ӧ����Ա����ӵ͵���
    Dim arrInfos() As Variant '����ǩ������Ϣ��
    Dim arrManIdxs() As Variant 'ǩ����Ա�����б��Index
    Dim arrSgnIdxs() As Variant 'ǩ����ť��Index
    '��ʼ��ǩ����ؽ���
    blnReadOnly = False: intCurr = -1: intHave = -1
    arrInfos = Array("סԺҽʦǩ��", "����ҽʦǩ��", "����ҽʦǩ��", "������ǩ��")
    arrManIdxs = Array(MC_סԺҽʦ, MC_����ҽʦ, MC_���λ�����, MC_������)
    arrSgnIdxs = Array(SL_סԺҽʦ, SL_����ҽʦ, SL_����ҽʦ, SL_������)

    On Error GoTo errH
    With gclsPros.CurrentForm
        For i = LBound(arrManIdxs) To UBound(arrManIdxs)
            .cboManInfo(arrManIdxs(i)).ForeColor = .ForeColor: .lblManInfo(arrManIdxs(i)).ForeColor = .ForeColor
            .cboManInfo(arrManIdxs(i)).Locked = False:  .cboManInfo(arrManIdxs(i)).BackColor = vbWindowBackground
            .cmdSign(arrSgnIdxs(i)).Caption = "ǩ��"
            If zlStr.NeedName(.cboManInfo(arrManIdxs(i)).Text) = UserInfo.���� Then
                intCurr = i
                .cmdSign(arrSgnIdxs(i)).Enabled = Not gclsPros.Is��ʿվ
            Else
                .cmdSign(arrSgnIdxs(i)).Enabled = False
            End If
            gclsPros.AuxiInfo.Filter = "��Ϣ��='" & arrInfos(i) & "'"
            If Not gclsPros.AuxiInfo.EOF Then
                intHave = i
                '��ǩ������ɫ�ֱ�ʾ
                .cboManInfo(arrManIdxs(i)).ForeColor = vbBlue: .lblManInfo(arrManIdxs(i)).ForeColor = vbBlue
                .cmdSign(arrSgnIdxs(i)).Caption = "ȡ��"
                'ǩ����ť�ɲ���״̬
                If gclsPros.AuxiInfo!��Ϣֵ & "" = UserInfo.���� Then
                    .cmdSign(arrSgnIdxs(i)).Enabled = Not gclsPros.Is��ʿվ
                Else '�����ѵ�ǩ������ȡ��
                    .cmdSign(arrSgnIdxs(i)).Enabled = False
                End If
            End If
        Next

        If intHave >= 0 Then
            '�漰ǩ������������ٸ���,��ȻȨ�޻���
            For i = LBound(arrManIdxs) To UBound(arrManIdxs)
                .cboManInfo(arrManIdxs(i)).Locked = True: .cboManInfo(arrManIdxs(i)).BackColor = vbButtonFace
                '�ͼ���ǩ�����ܱ��
                If i < intHave Then
                    .cmdSign(arrSgnIdxs(i)).Enabled = False
                End If
            Next
        End If

        '�����ǰ��Աǩ�����𲻸�����ǩ�������򲻿ɱ༭
        If intCurr <= intHave And intHave >= 0 Then
            blnReadOnly = True
        End If
    End With
    SetSignature = blnReadOnly

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CacheLoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String, Optional ByVal blnOnlyCache As Boolean, Optional ByVal intMaxDiagSource As Integer)
'���ܣ�����ϼ��ص�����в��һ���
'������vsDiagInput=��Ҫ������ϵı��
'      rsInput=��ȡ����ϼ�¼��
'      strDiagType=��������ַ������������Զ��ŷָ�
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���

    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, k As Long, LngRow As Long
    Dim bln�ֻ��̶� As Boolean
    Dim bln��ҽ As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnGet���� As Boolean

    blnGet���� = gclsPros.GetExtraCode
    On Error GoTo errH
    With vsDiagInput
        bln��ҽ = vsDiagInput.Name = "vsDiagXY"
        '�������
        If Not blnOnlyCache Then
            arrTmp = Split(strDiagType, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                Call FilterDiagByType(rsInput, Val(arrTmp(i)), intMaxDiagSource) '�������
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
                                                        If .TextMatrix(LngRow, DI_�������) <> "��Ժ���" And Val(.TextMatrix(LngRow, DI_��Ϸ���)) = 3 Then
                                .Cell(flexcpData, LngRow, DI_�������) = "�������"
                            End If
                        End If

                        If gclsPros.FuncType = f���ѡ�� Then
                            If InStr("," & gclsPros.DiagRowIDs & ",", "," & rsInput!ID & ",") > 0 Then
                                .TextMatrix(LngRow, DI_����) = 1
                            End If
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
                            .TextMatrix(LngRow, DI_��ϱ���) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!��ϱ��� & "")
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
                        If gclsPros.FuncType = f������ҳ Then
                            .TextMatrix(LngRow, DI_��ϱ���) = rsInput!�������� & ""
                            If Not IsNull(rsInput!����id) Then '�����ж���ϱ�����������Ƶ�һ����
                                .Cell(flexcpData, LngRow, DI_�������) = rsInput!�������� & ""
                                If Not gclsPros.CNIndent Or .TextMatrix(LngRow, DI_�������) = "" Then
                                    .TextMatrix(LngRow, DI_�������) = rsInput!�������� & ""
                                End If
                            End If
                        Else
                            If Not (IsNull(rsInput!���ID) And IsNull(rsInput!����id)) Then
                                .Cell(flexcpData, LngRow, DI_�������) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!������� & "")
                            Else
                                .Cell(flexcpData, LngRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                            End If
                        End If
                        If Val(rsInput!֤��ID & "") <> 0 And .TextMatrix(LngRow, DI_��ҽ֤��) = "" Then
                            .TextMatrix(LngRow, DI_��ҽ֤��) = rsInput!֤������ & ""
                        End If
                        .Cell(flexcpData, LngRow, DI_��ϱ���) = .TextMatrix(LngRow, DI_��ϱ���)
                        .Cell(flexcpData, LngRow, DI_��ҽ֤��) = .TextMatrix(LngRow, DI_��ҽ֤��)
                        If .TextMatrix(LngRow, DI_�������) <> "" Then
                            .AutoSize DI_��ϱ���, DI_�������
                        End If
                        If .ColWidth(DI_�������) < 3200 Then
                            .ColWidth(DI_�������) = 3200
                        End If
                        '���������ݼ�
                        .TextMatrix(LngRow, DI_����ʱ��) = Format(rsInput!����ʱ�� & "", "YYYY-MM-DD HH:mm")
                        .TextMatrix(LngRow, DI_��ע) = rsInput!��ע & ""
                        .TextMatrix(LngRow, DI_��Ժ���) = rsInput!��Ժ��� & ""
                        .TextMatrix(LngRow, DI_��Ժ����) = rsInput!��Ժ���� & ""
                        If blnGet���� Then
                            .TextMatrix(LngRow, DI_ICD����) = rsInput!���� & ""
                        End If
                        .TextMatrix(LngRow, DI_�Ƿ�δ��) = IIf(Val(rsInput!�Ƿ�δ�� & "") = 1, "��", "")
                        .TextMatrix(LngRow, DI_�Ƿ�����) = IIf(Val(rsInput!�Ƿ����� & "") = 1, "��", "")
                        If gclsPros.FuncType <> f������ҳ Then
                            .TextMatrix(LngRow, DI_���ID) = rsInput!���ID & ""
                        End If
                        .TextMatrix(LngRow, DI_����ID) = rsInput!����id & ""
                        .TextMatrix(LngRow, DI_֤��ID) = rsInput!֤��ID & ""
                        .TextMatrix(LngRow, DI_ҽ��IDs) = rsInput!ҽ��ID & ""
                        If gclsPros.FuncType = f������ҳ Then
                            If (arrTmp(i) = DT_��Ժ���XY Or arrTmp(i) = DT_��Ժ���ZY Or arrTmp(i) = DT_Ժ�ڸ�Ⱦ Or arrTmp(i) = DT_����֢) Then
'                                .TextMatrix(LngRow, DI_�̶�����) = IIf(IsNull(rsInput!����), "", "1")
                                .TextMatrix(LngRow, DI_�Ƿ���) = IIf(Val(rsInput!�Ƿ��� & "") = 1, "1", "")
                            End If
                        End If
                        .TextMatrix(LngRow, DI_��Ч����) = rsInput!��Ч���� & ""
                        .TextMatrix(LngRow, DI_������Ϣ) = IIf(IsNull(rsInput!����), "0", "1")
                        .TextMatrix(LngRow, DI_�����Դ) = Val(rsInput!��¼��Դ & "") '�����¼��Դ���Ա㱣��ʱ������Ϊ��ҳ�򲡰���Դ
                        .TextMatrix(LngRow, DI_��������) = rsInput!�������� & ""
                        .TextMatrix(LngRow, DI_�������) = rsInput!������� & ""
                        .TextMatrix(LngRow, DI_֤�����) = rsInput!֤����� & ""
                        .TextMatrix(LngRow, DI_��¼����) = Format(rsInput!��¼���� & "", "YYYY-MM-DD HH:mm")
                        .TextMatrix(LngRow, DI_��¼��Ա) = rsInput!��¼�� & ""
                        .RowData(LngRow) = Val(rsInput!ID & "")
                    Else
                        .TextMatrix(LngRow, DI_����ID) = rsInput!����id & ""
                        .TextMatrix(LngRow, DI_ICD����) = rsInput!�������� & ""
                        .Cell(flexcpData, LngRow, DI_ICD����) = .TextMatrix(LngRow, DI_ICD����)
                    End If
                    rsInput.MoveNext
                Loop
            Next
            '������������Ϣ
            If gclsPros.FuncType = f������ҳ Then Call SetDeliceryInfo(vsDiagInput)
        End If

        '���ݻ���
        strTmp = ""
        arrMain = Array(DI_��ϱ���, DI_��Ϸ���, DI_���ID, DI_����ID, DI_֤��ID, DI_��ҽ֤��)
        arrWhole = Array(DI_��Ϸ���, DI_��������, DI_��ϱ���, DI_ICD����, DI_�������, DI_֤�����, DI_��ҽ֤��, DI_�Ƿ�����, DI_֤��ID, DI_���ID, DI_����ID, DI_�������, DI_��ע, DI_����ʱ��, DI_��Ժ����, DI_��Ժ���, DI_�Ƿ�δ��)
        For i = .FixedRows To .Rows - 1
            blnFreeDiag = Val(.TextMatrix(i, DI_���ID)) = 0 And Val(.TextMatrix(i, DI_����ID)) = 0 '����¼�����
            If .TextMatrix(i, DI_�������) <> "" Then
                If strTmp <> .TextMatrix(i, DI_��Ϸ���) Then
                    j = 1: strTmp = .TextMatrix(i, DI_��Ϸ���)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = j
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    If strMainInfo = "" Then
                        strMainInfo = .TextMatrix(i, arrMain(k))
                    Else
                        strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                    End If
                Next
                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_�������) '����¼����ϼ����������
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), IIf(bln��ҽ, "��ҽ���", "��ҽ���"), strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, DI_�����Դ) = "", IIf(gclsPros.FuncType = f������ҳ, "4", "3"), .TextMatrix(i, DI_�����Դ)))
                '��ֹ���α��棬�޸���Դ
                If blnOnlyCache Then .TextMatrix(i, DI_�����Դ) = IIf(gclsPros.FuncType = f������ҳ, "4", "3")
            End If
        Next
        '���ظ���ID
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, DI_ICD����)) <> "" And Trim(.TextMatrix(i, DI_����ID)) = "" Then
                Set rsTmp = GetDiagExtraID(Trim(.TextMatrix(i, DI_ICD����)))
                If rsTmp.RecordCount > 0 Then
                    .TextMatrix(LngRow, DI_����ID) = rsTmp!ID & ""
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CacheLoadVsAllerData(ByRef vsAllerInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ���������Ϣ���ص�����в��һ���
'������vsAllerInput=��Ҫ���ع�����Ϣ�ı��
'      rsInput=������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    On Error GoTo errH
    With vsAllerInput
        If Not blnOnlyCache Then
            .Rows = .FixedRows
            For i = 1 To rsInput.RecordCount
                '������Դ�Ŀ������ظ�
                LngRow = -1
                If Not IsNull(rsInput!ҩ��ID) Then
                    LngRow = .FindRow(rsInput!ҩ��ID & "", , AI_ҩ��ID, , True)
                ElseIf Not IsNull(rsInput!ҩ����) Then
                    LngRow = .FindRow(rsInput!ҩ���� & "", , AI_����ҩ��, , True)
                End If
                If LngRow = -1 Then
                    For j = .FixedRows To .Rows - 1
                        If .TextMatrix(j, AI_����ҩ��) = "" Then
                            LngRow = j
                        End If
                    Next
                    If LngRow = -1 Then .Rows = .Rows + 1: LngRow = .Rows - 1
                    .TextMatrix(LngRow, AI_����ʱ��) = Format(rsInput!����ʱ��, "yyyy-MM-dd")
                    .TextMatrix(LngRow, AI_����ҩ��) = NVL(rsInput!ҩ����)
                    .TextMatrix(LngRow, AI_������Ӧ) = NVL(rsInput!������Ӧ)
                    .TextMatrix(LngRow, AI_����Դ����) = NVL(rsInput!����Դ����)
                    .TextMatrix(LngRow, AI_ҩ��ID) = rsInput!ҩ��ID & ""
                    .TextMatrix(LngRow, AI_������Դ) = rsInput!��¼��Դ & ""
                    '���ݱ��ݴ洢
                    .Cell(flexcpData, LngRow, AI_����ʱ��) = .TextMatrix(LngRow, AI_����ʱ��)
                    .Cell(flexcpData, LngRow, AI_����ҩ��) = .TextMatrix(LngRow, AI_����ҩ��)
                    .Cell(flexcpData, LngRow, AI_������Ӧ) = .TextMatrix(LngRow, AI_������Ӧ)
                    .Cell(flexcpData, LngRow, AI_����Դ����) = .TextMatrix(LngRow, AI_����Դ����)
                    .Cell(flexcpData, LngRow, AI_ҩ��ID) = .TextMatrix(LngRow, AI_ҩ��ID)
                    .RowData(LngRow) = Val(rsInput!ID & "")
                End If
                rsInput.MoveNext
            Next
            .Rows = .Rows + 1 '����һ�п���
            .Row = .FixedRows: .Col = AI_����ҩ��
        End If

        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, AI_����ҩ��) <> "" Then
                strInfo = .TextMatrix(i, AI_����ʱ��) & "|" & .TextMatrix(i, AI_����ҩ��) & "|" & .TextMatrix(i, AI_������Ӧ) & "|" & .TextMatrix(i, AI_����Դ����) & "|" & .TextMatrix(i, AI_ҩ��ID) & "|" & .RowData(i)
                strMainInfo = .TextMatrix(i, AI_����ʱ��) & "|" & .TextMatrix(i, AI_����Դ����) & "|" & .TextMatrix(i, AI_ҩ��ID) & "|" & .TextMatrix(i, AI_����ҩ��)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "����ҩ��", strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, AI_������Դ) = "", IIf(gclsPros.FuncType = f������ҳ, "4", "3"), .TextMatrix(i, AI_������Դ)))
                '��ֹ���α��棬�޸���Դ
                If blnOnlyCache Then .TextMatrix(i, AI_������Դ) = IIf(gclsPros.FuncType = f������ҳ, "4", "3")
            End If
        Next
    End With
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CacheLoadVsOPSData(ByRef vsOPSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ز����������ݲ�����
'������vsOPSInput=��Ҫ���ز���������Ϣ�ı��
'      rsInput=����������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    Dim lngOrder As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    With vsOPSInput
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, PI_��������) = Format(NVL(rsInput!������ʼʱ��, rsInput!��������) & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_��������) = Format(NVL(rsInput!��������ʱ��, rsInput!��������) & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_��������) = rsInput!�������� & ""
                .TextMatrix(i, PI_��������) = rsInput!�������� & ""
                If (Not gclsPros.CNIndent And gclsPros.FuncType = f������ҳ) Or .TextMatrix(i, PI_��������) = "" Then
                    .TextMatrix(i, PI_��������) = rsInput!����ԭ�� & ""
                    If .TextMatrix(i, PI_��������) = "" Then
                        .TextMatrix(i, PI_��������) = rsInput!�������� & ""
                    End If
                End If
                If .TextMatrix(i, PI_��������) <> "" Then
                    .AutoSize PI_��������, PI_��������
                End If
                .TextMatrix(i, PI_����ҽʦ) = rsInput!����ҽʦ & ""
                .TextMatrix(i, PI_������ʿ) = rsInput!������ʿ & ""
                .TextMatrix(i, PI_����1) = rsInput!��һ���� & ""
                .TextMatrix(i, PI_����2) = rsInput!�ڶ����� & ""
                .TextMatrix(i, PI_����ʽ) = rsInput!����ʽ & ""
                .TextMatrix(i, PI_����ҽʦ) = rsInput!����ҽʦ & ""
                If rsInput!�п� & rsInput!���� & "" <> "" Then
                    .TextMatrix(i, PI_�п�����) = rsInput!�п� & "/" & rsInput!����
                End If
                .TextMatrix(i, PI_��������ID) = rsInput!��������ID & ""
                .TextMatrix(i, PI_������ĿID) = rsInput!������Ŀid & ""
                .TextMatrix(i, PI_����ID) = rsInput!����ID & ""
                .TextMatrix(i, PI_��������) = rsInput!�������� & ""
                .TextMatrix(i, PI_�������) = rsInput!������� & ""
                .TextMatrix(i, PI_ASA�ּ�) = rsInput!asa�ּ� & ""
                .TextMatrix(i, PI_NNIS�ּ�) = rsInput!NNIS�ּ� & ""
                .TextMatrix(i, PI_��������) = rsInput!�������� & ""
                .TextMatrix(i, PI_�ٴ�����) = IIf(Val(rsInput!�ٴ����� & "") = 1, -1, 0)
                .TextMatrix(i, PI_׼������) = IIf(Val(rsInput!׼������ & "") = 0, "", Val(rsInput!׼������ & ""))
                .TextMatrix(i, PI_������ҩʱ��) = Format(rsInput!������ҩʱ�� & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_����ʼʱ��) = Format(rsInput!����ʼʱ�� & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_�пڲ�λ) = rsInput!�пڲ�λ & ""
                .TextMatrix(i, PI_�ط�������Ŀ��) = rsInput!�ط�Ŀ�� & ""
                .Cell(flexcpChecked, i, PI_�ط������Ҽƻ�) = Val(rsInput!�ط��ƻ� & "")
                .Cell(flexcpChecked, i, PI_�пڸ�Ⱦ) = Val(rsInput!�пڸ�Ⱦ & "")
                .Cell(flexcpChecked, i, PI_����֢) = Val(rsInput!����֢ & "")
                '10.34.10����
                .TextMatrix(i, PI_����ҩ����) = IIf(Val(rsInput!������ҩ���� & "") = 0, "", Val(rsInput!������ҩ���� & ""))
                .Cell(flexcpChecked, i, PI_Ԥ���ÿ���ҩ) = Val(rsInput!��ǰ������ҩ & "")
                .Cell(flexcpChecked, i, PI_��Ԥ�ڵĶ�������) = Val(rsInput!��Ԥ�ڵĶ������� & "")
                .Cell(flexcpChecked, i, PI_������֢) = Val(rsInput!������֢ & "")
                .Cell(flexcpChecked, i, PI_������������) = Val(rsInput!������������ & "")
                .Cell(flexcpChecked, i, PI_��������֢) = Val(rsInput!��������֢ & "")
                .Cell(flexcpChecked, i, PI_�����Ѫ��Ѫ��) = Val(rsInput!�����Ѫ��Ѫ�� & "")
                .Cell(flexcpChecked, i, PI_�����˿��ѿ�) = Val(rsInput!�����˿��ѿ� & "")
                .Cell(flexcpChecked, i, PI_�������Ѫ˨) = Val(rsInput!�������Ѫ˨ & "")
                .Cell(flexcpChecked, i, PI_���������л����) = Val(rsInput!���������л���� & "")
                .Cell(flexcpChecked, i, PI_�������˥��) = Val(rsInput!�������˥�� & "")
                .Cell(flexcpChecked, i, PI_�����˨��) = Val(rsInput!�����˨�� & "")
                .Cell(flexcpChecked, i, PI_�����Ѫ֢) = Val(rsInput!�����Ѫ֢ & "")
                .Cell(flexcpChecked, i, PI_�����Źؽڹ���) = Val(rsInput!�����Źؽڹ��� & "")
                .Cell(flexcpData, i, PI_��������) = rsInput!����ԭ�� & ""
                .TextMatrix(i, PI_������Դ) = rsInput!��¼��Դ & ""
                .RowData(i) = Val(rsInput!ID & "")
                '��¼���ڱ༭�ָ�
                For j = 0 To .Cols - 1
                    If j = PI_�������� And .TextMatrix(i, PI_��������) <> "" Then
                        If .Cell(flexcpData, i, j) = "" Then
                            .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                        End If
                    Else
                        .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                    End If
                Next

                If Trim(.TextMatrix(i, PI_��������)) <> "" And rsInput!ԭ�������� & "" <> "" Then
                    .Cell(flexcpData, i, PI_��������) = 1
                End If
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, PI_��������) <> "" Then
                lngOrder = lngOrder + 1
                strInfo = .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_����ҽʦ) & "|" & .TextMatrix(i, PI_������ʿ) & "|" & _
                        .TextMatrix(i, PI_����1) & "|" & .TextMatrix(i, PI_����2) & "|" & .TextMatrix(i, PI_����ʽ) & "|" & .TextMatrix(i, PI_����ҽʦ) & "|" & .TextMatrix(i, PI_�п�����) & "|" & .TextMatrix(i, PI_��������ID) & "|" & _
                        .TextMatrix(i, PI_������ĿID) & "|" & .TextMatrix(i, PI_����ID) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_�������) & "|" & .TextMatrix(i, PI_ASA�ּ�) & "|" & .TextMatrix(i, PI_NNIS�ּ�) & "|" & _
                        .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_�ٴ�����) & "|" & .TextMatrix(i, PI_׼������) & "|" & .TextMatrix(i, PI_������ҩʱ��) & "|" & .TextMatrix(i, PI_����ʼʱ��) & "|" & .TextMatrix(i, PI_�пڲ�λ) & "|" & _
                        .TextMatrix(i, PI_�ط�������Ŀ��) & "|" & .Cell(flexcpChecked, i, PI_�ط������Ҽƻ�) & "|" & .TextMatrix(i, PI_�пڸ�Ⱦ) & "|" & .Cell(flexcpChecked, i, PI_����֢) & "|" & .Cell(flexcpChecked, i, PI_Ԥ���ÿ���ҩ) & "|" & _
                        .TextMatrix(i, PI_����ҩ����) & "|" & .Cell(flexcpChecked, i, PI_��Ԥ�ڵĶ�������) & "|" & .Cell(flexcpChecked, i, PI_������֢) & "|" & .Cell(flexcpChecked, i, PI_������������) & "|" & .Cell(flexcpChecked, i, PI_��������֢) & "|" & _
                        .Cell(flexcpChecked, i, PI_�����Ѫ��Ѫ��) & "|" & .Cell(flexcpChecked, i, PI_�����˿��ѿ�) & "|" & .Cell(flexcpChecked, i, PI_�������Ѫ˨) & "|" & .Cell(flexcpChecked, i, PI_���������л����) & "|" & .Cell(flexcpChecked, i, PI_�������˥��) & "|" & _
                        .Cell(flexcpChecked, i, PI_�����˨��) & "|" & .Cell(flexcpChecked, i, PI_�����Ѫ֢) & "|" & .Cell(flexcpChecked, i, PI_�����Źؽڹ���) & "|" & .RowData(i) & "|" & lngOrder
                If gclsPros.MedPageSandard = ST_��������׼ Then
                    strMainInfo = .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������ID) & "|" & .TextMatrix(i, PI_������ĿID) & "|" & .TextMatrix(i, PI_�пڲ�λ)
                Else
                    strMainInfo = .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������) & "|" & .TextMatrix(i, PI_��������ID) & "|" & .TextMatrix(i, PI_������ĿID)
                End If
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "�������", strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, PI_������Դ) = "", IIf(gclsPros.FuncType = f������ҳ, "4", "3"), .TextMatrix(i, PI_������Դ)))
                '��ֹ���α��棬�޸���Դ
                If blnOnlyCache Then .TextMatrix(i, PI_������Դ) = IIf(gclsPros.FuncType = f������ҳ, "4", "3")
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Private Sub CacheLoadVsChemothData(ByRef vsChemothInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ػ������ݲ�����
'������vsChemothInput=��Ҫ���ػ�����Ϣ�ı��
'      rsInput=������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsChemothInput
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!��� & "")
                .TextMatrix(i, CI_��ѧ���Ʊ���) = NVL(rsInput!������Ϣ)
                .TextMatrix(i, CI_��ʼ����) = Format(rsInput!��ʼ����, "yyyy-MM-dd")
                .TextMatrix(i, CI_��������) = Format(rsInput!��������, "yyyy-MM-dd")
                .TextMatrix(i, CI_�Ƴ���) = Format(Val(rsInput!�Ƴ��� & ""), "###;-###;;")
                .TextMatrix(i, CI_����) = Format(Val(rsInput!���� & ""), "###;-###;;")
                .TextMatrix(i, CI_���Ʒ���) = rsInput!���Ʒ��� & ""
                .TextMatrix(i, CI_����Ч��) = rsInput!����Ч�� & ""
                .TextMatrix(i, CI_����ID) = rsInput!����id & ""
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, CI_��ѧ���Ʊ���) <> "" Then
                strInfo = .TextMatrix(i, CI_��ѧ���Ʊ���) & "|" & .TextMatrix(i, CI_��ʼ����) & "|" & .TextMatrix(i, CI_��������) & "|" & .TextMatrix(i, CI_�Ƴ���) & "|" & .TextMatrix(i, CI_����) & "|" & .TextMatrix(i, CI_���Ʒ���) & "|" & .TextMatrix(i, CI_����Ч��) & "|" & .TextMatrix(i, CI_����ID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, CI_��ѧ���Ʊ���)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "�������Ƽ�¼", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsRadiothData(ByRef vsRadiothInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ط������ݲ�����
'������vsRadiothInput=��Ҫ���ط�����Ϣ�ı��
'      rsInput=������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsRadiothInput
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!��� & "")
                .TextMatrix(i, RI_�������Ʊ���) = NVL(rsInput!������Ϣ)
                .TextMatrix(i, RI_��ʼ����) = Format(rsInput!��ʼ����, "yyyy-MM-dd")
                .TextMatrix(i, RI_��������) = Format(rsInput!��������, "yyyy-MM-dd")
                .TextMatrix(i, RI_�������) = Format(Val(rsInput!������� & ""), "###;-###;;")
                .TextMatrix(i, RI_�ۼ���) = Format(Val(rsInput!�ۼ��� & ""), "###;-###;;")
                .TextMatrix(i, RI_��Ұ��λ) = rsInput!��Ұ��λ & ""
                .TextMatrix(i, RI_����Ч��) = rsInput!����Ч�� & ""
                .TextMatrix(i, RI_����ID) = rsInput!����id & ""
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, RI_�������Ʊ���) <> "" Then
                strInfo = .TextMatrix(i, RI_�������Ʊ���) & "|" & .TextMatrix(i, RI_��ʼ����) & "|" & .TextMatrix(i, RI_��������) & "|" & .TextMatrix(i, RI_�������) & "|" & .TextMatrix(i, RI_�ۼ���) & "|" & .TextMatrix(i, RI_��Ұ��λ) & "|" & .TextMatrix(i, RI_����Ч��) & "|" & .TextMatrix(i, RI_����ID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, RI_�������Ʊ���)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "�������Ƽ�¼", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsSpiritData(ByRef vsSpiritInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ؾ���ҩƷʹ�����ݲ�����
'������vsSpiritInput=��Ҫ���ؾ���ҩƷ��Ϣ�ı��
'      rsInput=����ҩƷ��Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsSpiritInput
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!��� & "")
                .TextMatrix(i, SI_ҩ������) = rsInput!ҩ������ & ""
                .TextMatrix(i, SI_�Ƴ�) = rsInput!�Ƴ� & ""
                .TextMatrix(i, SI_�������) = rsInput!������� & ""
                .TextMatrix(i, SI_���ⷴӦ) = rsInput!���ⷴӦ & ""
                .TextMatrix(i, SI_��Ч) = rsInput!��Ч & ""
                .TextMatrix(i, SI_ҩƷID) = rsInput!ҩƷID & ""
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, SI_ҩ������) <> "" Then
                strInfo = .TextMatrix(i, SI_ҩ������) & "|" & .TextMatrix(i, SI_�Ƴ�) & "|" & .TextMatrix(i, SI_�������) & "|" & .TextMatrix(i, SI_���ⷴӦ) & "|" & .TextMatrix(i, SI_��Ч) & "|" & .TextMatrix(i, SI_ҩƷID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, SI_ҩ������)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "������������", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsKSSData(ByRef vsKSSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ؿ���ҩʹ�����ݲ�����
'������vsKSSInput=��Ҫ���ؿ���ҩ��Ϣ�ı��
'      rsInput=����ҩ��Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsKSSInput
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then Exit Sub
            Do While Not rsInput.EOF
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_����ҩ����) = "" Or (.RowData(i) = Val(rsInput!ҩ��id & "") And .TextMatrix(i, KI_��ҩĿ��) = rsInput!��ҩĿ�� & "" And .TextMatrix(i, KI_ʹ�ý׶�) = rsInput!ʹ�ý׶� & "") Then
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
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, KI_����ҩ����) <> "" Then
                strInfo = .RowData(i) & "|" & .TextMatrix(i, KI_����ҩ����) & "|" & .TextMatrix(i, KI_��ҩĿ��) & "|" & .TextMatrix(i, KI_ʹ�ý׶�) & "|" & .TextMatrix(i, KI_ʹ������) & "|" & .Cell(flexcpChecked, i, KI_һ���п�Ԥ����) & "|" & .TextMatrix(i, KI_DDD��) & "|" & .TextMatrix(i, KI_������ҩ) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, KI_����ҩ����) & "|" & .TextMatrix(i, KI_��ҩĿ��) & "|" & .TextMatrix(i, KI_ʹ�ý׶�)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "���˿����ؼ�¼", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsFlxAddICUData(Optional ByRef vsFlxAddICUInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ�������֢�໤ʹ�����ݲ�����
'������vsFlxAddICUInput=��Ҫ������֢�໤��Ϣ�ı��
'      rsInput=��֢�໤��Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String
    Dim strList As String
    Dim blnLocked As Boolean

    If Not vsFlxAddICUInput Is Nothing Then
        With vsFlxAddICUInput
            '���ݼ���
            If Not blnOnlyCache Then
                If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
                .Rows = rsInput.RecordCount + 2 '�̶���+����
                For i = 1 To rsInput.RecordCount
                    .TextMatrix(i, UI_�໤������) = rsInput!�໤������ & ""
                    .TextMatrix(i, UI_����ʱ��) = rsInput!����ʱ�� & ""
                    .TextMatrix(i, UI_�˳�ʱ��) = rsInput!�˳�ʱ�� & ""
                    If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                        .TextMatrix(i, UI_���) = i
                        .Cell(flexcpChecked, i, UI_����ס�ƻ�) = Val(rsInput!����ס�ƻ� & "")
                        .TextMatrix(i, UI_����סԭ��) = rsInput!����סԭ�� & ""
                    Else
                        .TextMatrix(i, UI_���) = Val(rsInput!��� & "")
                    End If
                    .RowData(i) = Val(rsInput!��� & "")
                    rsInput.MoveNext
                Next
            End If
            '���ݻ���
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, UI_�໤������) <> "" Then
                    strList = strList & "|" & .TextMatrix(i, UI_���) & "-" & .TextMatrix(i, UI_�໤������)
                    strInfo = .TextMatrix(i, UI_���) & "|" & .TextMatrix(i, UI_�໤������) & "|" & .TextMatrix(i, UI_����ʱ��) & "|" & .TextMatrix(i, UI_�˳�ʱ��) & "|" & .Cell(flexcpChecked, i, UI_����ס�ƻ�) & "|" & .TextMatrix(i, UI_����סԭ��) & "|" & .RowData(i)
                    strMainInfo = .TextMatrix(i, UI_�໤������) & "|" & .RowData(i)
                    Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "������֢�໤���", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
                End If
            Next
            strList = Mid(strList, 2)
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                gclsPros.CurrentForm.vsICUInstruments.ColComboList(TI_ICU����) = strList
                gclsPros.CurrentForm.vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
            End If
        End With
    Else
        '���ϰ棬û�б��
        If Not rsInput Is Nothing Then
            rsInput.Sort = "���"
            If Not rsInput.EOF Then
                rsInput.MoveFirst
                For i = 0 To rsInput.Fields.Count - 1
                    Call SetCtrlValues(rsInput.Fields(i).Name, rsInput.Fields(i).Value & "")
                Next
            End If
        End If
        blnLocked = gclsPros.CurrentForm.txtInfo(GC_��֢�໤������).Text = ""
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�˹������ѳ�), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�ط���֢ҽѧ��), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�ط����ʱ��), blnLocked, True)
    End If
End Sub

Private Sub CacheLoadVsICUInstrumentsData(ByRef vsICUInstruments As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ�������е����ʹ��������ݲ�����
'������vsICUInstruments=��Ҫ������е����ʹ�������Ϣ�ı��
'      rsInput=��е����ʹ�������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsICUInstruments
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, TI_ICU����) = rsInput!�໤������ & ""
                .Cell(flexcpData, i, TI_ICU����) = Val(rsInput!��� & "")
                .TextMatrix(i, TI_��е������) = rsInput!��е������ & ""
                .TextMatrix(i, TI_��ʼʱ��) = rsInput!��ʼʹ��ʱ�� & ""
                .TextMatrix(i, TI_����ʱ��) = rsInput!����ʹ��ʱ�� & ""
                .TextMatrix(i, TI_��Ⱦ�ۼ�Сʱ) = rsInput!��Ⱦ�ۼ�ʱ�� & ""
                .RowData(i) = Val(rsInput!��� & "")
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, TI_��е������) <> "" Then
                strInfo = .TextMatrix(i, TI_ICU����) & "|" & .TextMatrix(i, TI_��е������) & "|" & .TextMatrix(i, TI_��ʼʱ��) & "|" & .TextMatrix(i, TI_����ʱ��) & "|" & .TextMatrix(i, TI_��Ⱦ�ۼ�Сʱ) & "|" & i
                strMainInfo = .TextMatrix(i, TI_ICU����) & "|" & .TextMatrix(i, TI_��е������) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "��е����ʹ�����", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadvsInfectData(ByRef vsInfect As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ز��˸�Ⱦ��¼���ݲ�����
'������vsInfect=��Ҫ���ز��˸�Ⱦ��¼��Ϣ�ı��
'      rsInput=���˸�Ⱦ��¼��Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsInfect
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, FI_ȷ������) = rsInput!ȷ������ & ""
                .TextMatrix(i, FI_��Ⱦ��λ) = rsInput!��Ⱦ��λ & ""
                .TextMatrix(i, FI_ҽԺ��Ⱦ����) = rsInput!ҽԺ��Ⱦ���� & ""
                .TextMatrix(i, FI_ҽԺ��Ⱦ����) = rsInput!ҽԺ��Ⱦ���� & ""
                .RowData(i) = Val(rsInput!��� & "")
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, FI_��Ⱦ��λ) <> "" Then
                strInfo = .TextMatrix(i, FI_ȷ������) & "|" & .TextMatrix(i, FI_��Ⱦ��λ) & "|" & .TextMatrix(i, FI_ҽԺ��Ⱦ����) & "|" & i
                strMainInfo = .TextMatrix(i, FI_��Ⱦ��λ) & "|" & .TextMatrix(i, FI_ҽԺ��Ⱦ����) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "���˸�Ⱦ��¼", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadvsSampleData(ByRef vsSample As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ����ز��˲�ԭѧ������ݲ�����
'������vsSample=��Ҫ���ز��˲�ԭѧ�����Ϣ�ı��
'      rsInput=���˲�ԭѧ�����Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsSample
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '�̶���+����
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, MI_�걾) = decode(Val(rsInput!�걾 & ""), 1, "1.ѪҺ", 2, "2.��Һ", 3, "3.���", 4, "4.̵Һ", 5, "5.����������")
                .TextMatrix(i, MI_��ԭѧ���뼰����) = rsInput!��ԭѧ���� & ""
                .TextMatrix(i, MI_�ͼ�����) = rsInput!�ͼ����� & ""
                .RowData(i) = Val(rsInput!��� & "")
                rsInput.MoveNext
            Next
        End If
        '���ݻ���
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, MI_�걾) <> "" Then
                strInfo = .TextMatrix(i, MI_�걾) & "|" & .TextMatrix(i, MI_��ԭѧ���뼰����) & "|" & .TextMatrix(i, MI_�ͼ�����) & "|" & i
                strMainInfo = .TextMatrix(i, MI_�걾) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "���˲�ԭѧ���", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsFreesData(ByRef vsFees As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean, Optional ByVal bln��Ŀ As Boolean)
'���ܣ����ز���סԺ�������ݲ�����
'������vsFees=��Ҫ���ز���סԺ���õı��
'      rsInput=����סԺ������Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'      bln��Ŀ=���ݼ���ʱ�����Ƿ������Ѿ���Ŀ��Ϣ
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngCol As Long, LngRow As Long, lng��� As Long
    Dim strInfo As String, strMainInfo As String
    Dim dblӤ���� As Double, dblSum As Double
    Dim blnHave As Boolean, blnӤ�� As Boolean
    On Error GoTo errH
    With vsFees
        '���ݼ���
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 2: Exit Sub
            If Not bln��Ŀ Then
                rsInput.Filter = "Ӥ����<>0"
                Do While Not rsInput.EOF
                    dblӤ���� = dblӤ���� + Val(rsInput!��� & "")
                    rsInput.MoveNext
                Loop
                rsInput.Filter = "Ӥ����=0"
            End If
            .Rows = .FixedRows + rsInput.RecordCount \ 3 + 1

            For i = 0 To rsInput.RecordCount - 1
                If i Mod 3 = 0 Then LngRow = LngRow + 1 '3��������λ����һ��
                LngCol = (i Mod 3) * 2 '��λ��
                blnӤ�� = rsInput!��Ŀ���� = "Ӥ����"
                If blnӤ�� Then blnHave = True
                .TextMatrix(LngRow, LngCol) = rsInput!����
                .TextMatrix(LngRow, LngCol + 1) = Format(Val(rsInput!��� & "") + IIf(blnӤ��, dblӤ����, 0), gclsPros.FreeFormat)
                rsInput.MoveNext
            Next

            If dblӤ���� <> 0 And Not blnHave Then
                If LngCol = 4 Then
                    LngCol = 0: LngRow = LngRow + 1 'д���ˣ��ƶ�����һ��
                Else
                    LngCol = LngCol + 2 '�ƶ�����һ��
                End If
                .TextMatrix(LngRow, LngCol) = "Ӥ����"
                .TextMatrix(LngRow, LngCol + 1) = Format(dblӤ����, gclsPros.FreeFormat)
                If LngCol = 4 Then .Rows = .Rows + 1 '���Ӥ����д���ǵ�����������������
            End If
            Call SumAndSetFrees
        End If
        '���ݻ���
        '����δ��Ŀ���ݣ���һ�β�����,��Щ���ݻᵱ����������
        If bln��Ŀ Or blnOnlyCache Then
            For i = 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) <> "" Then
                    strInfo = .TextMatrix(LngRow, LngCol)
                    strMainInfo = .TextMatrix(LngRow, LngCol) & "|" & .TextMatrix(LngRow, LngCol + 1)
                    Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "���˷���", strMainInfo, strInfo, , , LngRow & "," & LngCol)
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Sub LoadTransferData(Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean, Optional ByVal bln��Ŀ As Boolean)
'���ܣ����ز���סԺת����Ϣ������
'������
'      rsInput=����סԺת����Ϣ��¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'      bln��Ŀ=���ݼ���ʱ�Ƿ������Ѿ���Ŀ��Ϣ
'˵����LoadMedPageData���Ӻ���
    Dim i As Long, LngCol As Long, LngRow As Long
    Dim vsTranfer As VSFlexGrid

    Dim strInfo As String, strMainInfo As String
    If gclsPros.FuncType <> f������ҳ Then
        With gclsPros.CurrentForm
            If .txtInfo(GC_ת��1).Text = "" And .txtInfo(GC_ת��2).Text = "" And .txtInfo(GC_ת��3).Text = "" Then
                For i = 1 To rsInput.RecordCount
                    If i = 1 Then
                        .txtInfo(GC_ת��1).Text = rsInput!�������� & ""
                    ElseIf i = 2 Then
                        .txtInfo(GC_ת��2).Text = rsInput!�������� & ""
                    ElseIf i = 3 Then
                        .txtInfo(GC_ת��3).Text = rsInput!�������� & ""
                        Exit For
                    End If
                    rsInput.MoveNext
                Next
            End If
        End With
    Else
        Set vsTranfer = gclsPros.CurrentForm.vsTransfer
        With vsTranfer
            For i = 1 To rsInput.RecordCount
                .TextMatrix(0, i) = rsInput!��������
                .TextMatrix(1, i) = Format(rsInput!��ʼʱ��, "YYYY-MM-DD")
                If i = 6 Then Exit For
                rsInput.MoveNext
            Next
        End With
    End If
End Sub

Public Function FreeHaveLowLevel(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
'����:�жϷ��ü����Ƿ�����¼�����
    Dim strCode As String
    Dim lngPos As Long, i As Integer
    Dim vsTmp As VSFlexGrid

    Set vsTmp = gclsPros.CurrentForm.vsFees
    With vsTmp
        If .TextMatrix(LngRow, LngCol) = "" Then Exit Function
        strCode = GetFreeCode(.TextMatrix(LngRow, LngCol), True)
        If strCode = "" Then FreeHaveLowLevel = True: Exit Function
        lngPos = InStr(1, strCode, "_") '���ñ����ʽ����������_����
        '�ж��Ƿ�������һ�����ã���ȡ��ǰ����
        If lngPos > 0 Then strCode = Mid(strCode, lngPos + 1)

        For i = 3 To (.Rows * 3) - 1
            LngRow = i \ 3: LngCol = (i Mod 3) * 2 '��λ����
            If .TextMatrix(LngRow, LngCol) Like strCode & "_*.*" Then
                FreeHaveLowLevel = True: Exit Function  '�����Ӽ��ж��˳�
            End If
        Next
    End With
End Function

Public Sub SumAndSetFrees()
'����:�����������ۼƣ������õ�Ԫ���ֵ�Լ��ܷ���
    Dim strCode As String, strFathCode As String
    Dim dblSum As Double, lngPos As Long, i As Long, j As Long
    Dim LngRow As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid
    Dim intID As Integer
    Dim blnDo As Boolean, strFee As String

    Dim rsFeeList As New ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim lngSort As Long

    On Error GoTo errH
    rsFeeList.Fields.Append "ID", adInteger, , adFldKeyColumn           '����
    rsFeeList.Fields.Append "Row", adInteger                            '�к�
    rsFeeList.Fields.Append "Col", adInteger                            '�к�
    rsFeeList.Fields.Append "Code", adVarChar, 200                      '����
    rsFeeList.Fields.Append "PID", adInteger, , adFldIsNullable         '����ID
    rsFeeList.Fields.Append "Fee", adVarChar, 200, adFldIsNullable      '�����ַ���
    rsFeeList.Fields.Append "Sort", adInteger, , adFldIsNullable        '����
    rsFeeList.CursorLocation = adUseClient
    rsFeeList.LockType = adLockOptimistic
    rsFeeList.CursorType = adOpenStatic
    rsFeeList.Open

    Set vsTmp = gclsPros.CurrentForm.vsFees
    With vsTmp
        For i = 3 To (.Rows * 3) - 1
            LngRow = i \ 3: LngCol = (i Mod 3) * 2
            strCode = GetFreeCode(.TextMatrix(LngRow, LngCol))
            If strCode <> "" Then
                rsFeeList.AddNew Array("ID", "Row", "Col", "Code", "Sort"), Array(Identity(lngSort), LngRow, LngCol, strCode, 0)
            End If
        Next
        rsFeeList.Filter = "": rsFeeList.Sort = "Code,ID"
        Set rsTmp = zlDatabase.CopyNewRec(rsFeeList) '���ݼ�¼��������������ID
        rsTmp.Filter = "": rsTmp.Sort = "Code,ID"
        For i = 1 To rsTmp.RecordCount
            strCode = rsTmp!Code & ""
            lngPos = InStr(strCode, "_")
            If lngPos > 0 Then '��ȡ��ǰ������Ϊ������
                strFathCode = Mid(strCode, lngPos + 1)
            Else
                strFathCode = strCode
            End If
            '�����Ӽ�
            Call Rec.Update(rsFeeList, "Code Like '" & strFathCode & "_*'", "PID", rsTmp!ID)
            rsTmp.MoveNext
        Next
        '���á������������֮��
        rsFeeList.Filter = "": rsFeeList.Sort = "ID"
        Set rsTmp = zlDatabase.CopyNewRec(rsFeeList) '���ݼ�¼��������������ID
        Do While JudeSet(rsFeeList)
            intID = Val(rsFeeList!ID & "")
            rsTmp.Filter = "PID=" & intID
            blnDo = False
            If rsTmp.EOF Then '��ǰ��������׼�
                blnDo = True
            Else '��ǰ���ò�����ͼ�
                rsTmp.Filter = "PID=" & intID & " And Fee=Null"
                dblSum = 0
                If rsTmp.EOF Then '���õ��Ӽ�����ȡ����,��ǰ�����Ӽ��������
                    rsTmp.Filter = "PID=" & intID
                    Do While Not rsTmp.EOF
                        dblSum = dblSum + Val(rsTmp!Fee & "")
                        rsTmp.MoveNext
                    Loop
                    .TextMatrix(rsFeeList!Row, rsFeeList!Col + 1) = Format(dblSum, gclsPros.FreeFormat) '�����������ڱ����
                    blnDo = True
                End If
            End If
            If blnDo Then '����������д�ڼ�¼����
                rsTmp.Filter = "ID=" & intID
                strFee = Format(Val(.TextMatrix(rsFeeList!Row, rsFeeList!Col + 1)), gclsPros.FreeFormat)
                rsFeeList.Update "Fee", strFee
                rsTmp.Update "Fee", strFee
            Else
                rsFeeList.Update "Sort", Val(rsFeeList!Sort & "") + 1 '�����ֶ����ӣ�˳�����ں���
            End If
        Loop
        '���㲢�����ܷ���
        rsFeeList.Filter = "PID=Null"
        dblSum = 0
        Do While Not rsFeeList.EOF
            dblSum = dblSum + Val(rsFeeList!Fee & "")
            rsFeeList.MoveNext
        Loop
        gclsPros.CurrentForm.txtSpecificInfo(SLC_���ú�).Text = Format(dblSum, gclsPros.FreeFormat)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function JudeSet(ByRef rsInput As ADODB.Recordset) As Boolean
'���ܣ��ж��Ƿ���δ��ȡ���ķ���
    rsInput.Filter = "Fee=Null"
    rsInput.Sort = "Sort,ID"
    JudeSet = Not rsInput.EOF
End Function
Public Function GetFreeCode(ByVal strFreeName As String, Optional ByVal blnMustHave As Boolean) As String
'���ܣ���ȡ���ñ���
'������strFreeName=������
'      blnMustHave=�Ƿ�����б��룬û�б��뷵�ؿ�
'���أ����ñ���
'˵����SetHlevelFreeSum���Ӻ���
    Dim strCode As String
    Dim lngPos As Long

    lngPos = InStr(strFreeName, ".")
    If lngPos > 0 Then strCode = Mid(strFreeName, 1, lngPos - 1)
    If blnMustHave And lngPos <= 0 Then strCode = ""
    GetFreeCode = strCode
End Function

Public Sub FilterDiagByType(ByRef rsInput As ADODB.Recordset, ByVal intDiagType As Integer, Optional ByVal intMaxDiagSource As Integer = -1)
'���ܣ���ȡ�ƶ����͵����
'������rsInput=��Ҫ���˵���ϼ�¼��
'      intDiagType-�������
'      intMaxDiagSource=���ļ�¼��Դ��������ҳʹ��,���������ȡ��ҳʱ���������Դ�Ӵ�С������ȡ
'���أ�rsInput=�����˵���ϼ�¼��
'˵����LoadMedPageData���Ӻ���
    Dim blnDo As Boolean
    '�ǲ�����ҳ��ÿһ��������Դ���ȼ���ȡ�������Ҫ������ж��Ƿ�ȫ������
    Select Case intMaxDiagSource
        Case -1, 1, 2
            blnDo = True
        Case Else
            rsInput.Filter = "��¼��Դ=" & intMaxDiagSource & " And �������=" & intDiagType
    End Select

    If blnDo Then
        If intMaxDiagSource > 0 Then '������ҳ
            If Val(intDiagType) <> 21 Then
                If rsInput.EOF Then
                    rsInput.Filter = "��¼��Դ=2 And �������=" & intDiagType
                End If
                If rsInput.EOF Then
                    rsInput.Filter = "��¼��Դ=1 And �������=" & intDiagType
                End If
            End If
        Else 'סԺ��ҳ
            rsInput.Filter = "��¼��Դ=3 And �������=" & intDiagType
            If Val(intDiagType) <> 21 Then
                If rsInput.EOF Then
                    rsInput.Filter = "��¼��Դ=2 And �������=" & intDiagType
                End If
                If rsInput.EOF Then
                    rsInput.Filter = "��¼��Դ=1 And �������=" & intDiagType
                End If
            End If
            If rsInput.EOF Then
                rsInput.Filter = "��¼��Դ=4 And �������=" & intDiagType
            End If
        End If
    End If
End Sub

Public Function FindDiagRow(ByVal dtInput As DiagType) As Long
'���ܣ���ȡָ��������ϵ��к�
'������dtInput=�������
    Dim bln��ҽDiag  As Boolean
    Dim i As Long, LngRow As Long

    If dtInput <= DT_��Ժ���ZY And dtInput > DT_����֢ Then bln��ҽDiag = True
    With IIf(bln��ҽDiag, gclsPros.CurrentForm.vsDiagZY, gclsPros.CurrentForm.vsDiagXY)
        LngRow = .FindRow(dtInput & "", , DI_��Ϸ���)
        FindDiagRow = LngRow
    End With
End Function

Public Sub CacheLoadDiagMatchData(Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'���ܣ�������Ϸ���������ݲ�����
'������rsInput=��Ϸ��������¼��
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
    Dim arrCtrlIdxs() As Variant
    Dim arrInfoIdxs() As Variant
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim strTmp As String

    On Error GoTo errH
    If gclsPros.FuncType = f���ѡ�� And gclsPros.PatiType = PF_סԺ Then
        If gclsPros.IsTCM Then
            arrCtrlIdxs = Array(BCC_�������ԺXY, BCC_��Ժ���ԺXY, BCC_�����벡��, BCC_�ٴ��벡��, BCC_�ٴ���ʬ��, BCC_��������Ժ, _
                            BCC_�������ԺZY, BCC_��Ժ���ԺZY)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 7, 11, 12)
        Else
            arrCtrlIdxs = Array(BCC_�������ԺXY, BCC_��Ժ���ԺXY, BCC_�����벡��, BCC_�ٴ��벡��, BCC_�ٴ���ʬ��, BCC_��������Ժ)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 7)
        End If
    Else
        If gclsPros.IsTCM Then
            arrCtrlIdxs = Array(BCC_�������ԺXY, BCC_��Ժ���ԺXY, BCC_�����벡��, BCC_�ٴ��벡��, BCC_�ٴ���ʬ��, BCC_��ǰ������, BCC_��������Ժ, _
                            BCC_�������ԺZY, BCC_��Ժ���ԺZY, BCC_��֤, BCC_�η�, BCC_��ҩ)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15)
        Else
            arrCtrlIdxs = Array(BCC_�������ԺXY, BCC_��Ժ���ԺXY, BCC_�����벡��, BCC_�ٴ��벡��, BCC_�ٴ���ʬ��, BCC_��ǰ������, BCC_��������Ժ)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 6, 7)
        End If
    End If
    If Not blnOnlyCache Then
        For i = LBound(arrCtrlIdxs) To UBound(arrCtrlIdxs)
            '������Ϸ������ȱʡֵ
            Call SetDiagMatchInfo(arrCtrlIdxs(i))
        Next
    End If
    For i = LBound(arrCtrlIdxs) To UBound(arrCtrlIdxs)
        Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrCtrlIdxs(i))
        If Not blnOnlyCache Then
            If Not rsInput Is Nothing Then
                '������Ϸ������
                strTmp = ""
                rsInput.Filter = "��������=" & arrInfoIdxs(i)
                If Not rsInput.EOF Then
                    If Val(rsInput!������� & "") >= 0 Then
                        Call zlControl.CboSetIndex(objCboTmp.hwnd, rsInput!�������)
                        strTmp = IIf(Not objCboTmp.Locked, rsInput!������� & "", "")
                    End If
                End If
                Call UpdateCacheRecInfo(0, "��Ϸ������", strTmp, strTmp, arrCtrlIdxs(i))
            End If
        Else
            strTmp = ""
            If Not objCboTmp.Locked Then
                strTmp = IIf(objCboTmp.ListIndex = -1, "", objCboTmp.ListIndex)
            Else
                If arrCtrlIdxs(i) = BCC_�ٴ���ʬ�� Then
                    strTmp = IIf(objCboTmp.ListIndex = -1, "", 4)
                End If
            End If
            Call UpdateCacheRecInfo(1, "��Ϸ������", strTmp, strTmp, arrCtrlIdxs(i))
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CompareDiag(ByVal strTmp1 As String, ByVal strTmp2 As String) As Boolean
'���ܣ���������Ͻ��жԱȣ���һ���������ͬ�ľͷ��� true
    Dim arrTmp1() As String
    Dim arrTmp2() As String
    Dim i As Long, j As Long

    arrTmp1 = Split(strTmp1, Chr(10))
    arrTmp2 = Split(strTmp2, Chr(10))

    For i = LBound(arrTmp1) To UBound(arrTmp1)
        For j = LBound(arrTmp2) To UBound(arrTmp2)
            If arrTmp1(i) = arrTmp2(j) Then
                CompareDiag = True
                Exit Function
            End If
        Next
    Next
End Function

Private Function GetStrDiagS(ByVal intIdx As Integer) As String
'���ܣ���ȡĳһ���͵�������ϣ���Chr(10)���зָ�
'������intIdx=�������
    Dim i As Long
    Dim lngRow1 As Long, lngRow2 As Long
    Dim strTmp1 As String, strTmp2 As String
    Dim bln��ҽ As Boolean
    Dim vsDiag As VSFlexGrid

    If intIdx = DT_�������XY Then
        lngRow1 = FindDiagRow(DT_�������XY)
        lngRow2 = FindDiagRow(DT_��Ժ���XY) - 1
    ElseIf intIdx = DT_��Ժ���XY Then
        lngRow1 = FindDiagRow(DT_��Ժ���XY)
        lngRow2 = FindDiagRow(DT_��Ժ���XY) - 1
    ElseIf intIdx = DT_��Ժ���XY Then
        lngRow1 = FindDiagRow(DT_��Ժ���XY)
        lngRow2 = FindDiagRow(DT_Ժ�ڸ�Ⱦ) - 1
    ElseIf intIdx = DT_�������ZY Then
        bln��ҽ = True
        lngRow1 = FindDiagRow(DT_�������ZY)
        lngRow2 = FindDiagRow(DT_��Ժ���ZY) - 1
    ElseIf intIdx = DT_��Ժ���ZY Then
        bln��ҽ = True
        lngRow1 = FindDiagRow(DT_��Ժ���ZY)
        lngRow2 = FindDiagRow(DT_��Ժ���ZY) - 1
    ElseIf intIdx = DT_��Ժ���ZY Then
        bln��ҽ = True
        lngRow1 = FindDiagRow(DT_��Ժ���ZY)
        lngRow2 = gclsPros.CurrentForm.vsDiagZY.Rows - 1
    End If

    Set vsDiag = IIf(bln��ҽ, gclsPros.CurrentForm.vsDiagZY, gclsPros.CurrentForm.vsDiagXY)
    For i = lngRow1 To lngRow2
        If Trim(vsDiag.TextMatrix(i, DI_�������)) <> "" Then
            strTmp1 = strTmp1 & Chr(10) & Trim(vsDiag.TextMatrix(i, DI_�������))
            strTmp2 = strTmp2 & Trim(vsDiag.TextMatrix(i, DI_�������))
        End If
    Next

    If strTmp2 = "" Then
        strTmp1 = ""
    Else
        strTmp1 = Mid(strTmp1, InStr(strTmp1, Chr(10)) + 1)
    End If
    GetStrDiagS = strTmp1
End Function

Public Sub SetDiagMatchInfo(ByVal intIdx As Integer, Optional ByVal blnJustState As Boolean)
'���ܣ�����Ϸ����������ȱʡֵ�����Լ�����Ƿ��������
'������intIdx=Ҫ���õķ�������ؼ�
'      blnJustState=ֻ���÷������״̬
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim strTmp1 As String, strTmp2 As String
    Dim lngHwnd As Long, lngIndex As Long
    Dim blnNotOther As Boolean
    Dim blnMedRecChange As Boolean

    With gclsPros.CurrentForm
        Set objCboTmp = .cboBaseInfo(intIdx)
        blnNotOther = True: lngIndex = -1
         '�������Ժ��������Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = BCC_�������ԺXY Then
            strTmp1 = GetStrDiagS(DT_�������XY)
            strTmp2 = GetStrDiagS(DT_��Ժ���XY)
        '��Ժ���Ժ����Ժ��Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        ElseIf intIdx = BCC_��Ժ���ԺXY Then
            strTmp1 = GetStrDiagS(DT_��Ժ���XY)
            strTmp2 = GetStrDiagS(DT_��Ժ���XY)
        '��������Ժ��������Ϻ���Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        ElseIf intIdx = BCC_��������Ժ Then
            strTmp1 = GetStrDiagS(DT_�������XY)
            strTmp2 = GetStrDiagS(DT_��Ժ���XY)
        '��ҽ�������Ժ��������Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        ElseIf intIdx = BCC_�������ԺZY Then
            strTmp1 = GetStrDiagS(DT_�������ZY)
            strTmp2 = GetStrDiagS(DT_��Ժ���ZY)
        '��ҽ��Ժ���Ժ����Ժ��Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        ElseIf intIdx = BCC_��Ժ���ԺZY Then
            strTmp1 = GetStrDiagS(DT_��Ժ���ZY)
            strTmp2 = GetStrDiagS(DT_��Ժ���ZY)
        Else
           blnNotOther = False
        End If

        If blnNotOther Then
            If strTmp1 & strTmp2 = "" Then
                lngIndex = 0
            ElseIf strTmp1 = "" Or strTmp2 = "" Then
                lngIndex = 3
            Else
                lngIndex = IIf(CompareDiag(strTmp1, strTmp2), 1, 2)
            End If
        Else
            '�����벡���ٴ��벡��¼�벡����Ϻ����¼�룬ȱʡΪ���ϡ�
            If intIdx = BCC_�����벡�� Or intIdx = BCC_�ٴ��벡�� Then
                strTmp1 = .vsDiagXY.TextMatrix(FindDiagRow(DT_�������), DI_�������)
                Call SetCtrlLocked(objCboTmp, strTmp1 = "")
                If strTmp1 <> "" Then
                    lngIndex = 1
                    objCboTmp.BackColor = vbWindowBackground
                Else
                    lngIndex = 0
                    objCboTmp.BackColor = vbButtonFace
                End If
            '�ٴ���ʬ�죺��ѡʬ������¼�룬ȱʡΪ���ϡ�
            ElseIf intIdx = BCC_�ٴ���ʬ�� Then
                Call SetCtrlLocked(objCboTmp, .cboBaseInfo(BCC_��������ʬ��).ListIndex <= 0)
                If .cboBaseInfo(BCC_��������ʬ��).ListIndex = 1 Then
                    lngIndex = 1
                    objCboTmp.BackColor = vbWindowBackground
                Else
                    lngIndex = 0
                    objCboTmp.BackColor = vbButtonFace
                End If
                blnMedRecChange = True
            '��ǰ����������������������¼�룬ȱʡΪ���ϡ�
            ElseIf intIdx = BCC_��ǰ������ Then
                For i = .vsOPS.FixedRows To .vsOPS.Rows - 1
                    If Trim(.vsOPS.TextMatrix(i, PI_��������)) <> "" Then Exit For
                Next
                If i > .vsOPS.Rows - 1 Then
                    lngIndex = 0                '�����Ը�ʱȱʡΪδ��
                Else
                    lngIndex = 1
                End If
                blnMedRecChange = True
            End If
            If blnJustState Then
                lngIndex = objCboTmp.ListIndex
            End If
        End If
        
        If lngIndex > -1 Then
            If gclsPros.FuncType = f������ҳ Then
                If blnMedRecChange Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, lngIndex)
                End If
            Else
                Call zlControl.CboSetIndex(objCboTmp.hwnd, lngIndex)
            End If
        End If
    End With
End Sub

Public Sub UpdateCacheRecInfo(Optional ByRef intType As Integer, Optional ByVal strInfoName As String, Optional ByVal strWholeInfo As String, Optional ByVal strMainInfo As String, Optional ByVal lngRowNo As Long = -1, Optional ByVal lngID As Long, Optional ByRef strTag As String)
'���ܣ����»������Ϣ��¼����һ��Ӧ���ڱ��
'������intType=0-��ʼ�����¼��أ�1-����ǰ���ͨ����ļ��ػ����,2-������Ϣ��¼�����£����ڼ��ؼ��ĸı�״̬
'      strInfoName=��Ϣ����ؼ���
'      strWholeInfo=��Ϣ����
'      strMainInfo=����Ϣֵ
'      lngRowNo=�кŻ�ؼ����
'      lngId=ID��
'      strTag=������λ��Ϣ,intType=0ʱ��д��intType=1ʱ����
'ע�������չ��Ϣ���μ���Ϣ��¼�������ݵģ�strWholeInfo��strMainInfo��Ҫ�������߿�����ͬ
    Dim lngSort As Long, lng״̬ As Long
    Dim lng��� As Long
    Dim i As Long
    Dim rsTmp As ADODB.Recordset, strFilter As String
    Dim strTmp As String, arrTmp As Variant

    On Error GoTo errH
    strWholeInfo = Trim(strWholeInfo)
    strMainInfo = Trim(strMainInfo)
    strTag = Trim(strTag)
    If intType <> 2 Then
        '��������Ϣ��Ѱ��Ѱ�ң�Ѱ�Ҳ���ʱ���ٰ��ؼ���Ѱ��
        gclsPros.MainInfoRec.Filter = "��Ϣ��='" & strInfoName & "'"
        If gclsPros.MainInfoRec.EOF Then gclsPros.MainInfoRec.Filter = "�ؼ���='" & strInfoName & "'" & IIf(lngRowNo = -1, "", " And Index=" & lngRowNo)
        If Not gclsPros.MainInfoRec.EOF Then
            Select Case gclsPros.MainInfoRec!ExpState
                Case ES_������չ
                    Call gclsPros.MainInfoRec.Update(IIf(intType = 0, "��Ϣԭֵ", "��Ϣ��ֵ"), strWholeInfo)
                Case ES_��ʼ��չ
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And IndexEx=" & lngRowNo & IIf(strTag <> "" And intType = 1, " And Tag='" & strTag & "'", "")
                    If intType = 0 Then
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣԭֵ", "����Ϣԭֵ", "Tag"), Array(strWholeInfo, strMainInfo, strTag))
                    Else
                        Call gclsPros.SecdInfoRec.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ", "Tag"), Array(strWholeInfo, strMainInfo, strTag))
                    End If
                Case ES_������չ
                    gclsPros.SecdInfoRec.Filter = "": gclsPros.SecdInfoRec.Sort = "Sort"
                    If Not gclsPros.SecdInfoRec.EOF Then gclsPros.SecdInfoRec.MoveLast: lngSort = gclsPros.SecdInfoRec!Sort
                    If intType = 0 Then
                        gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And IndexEx=" & lngRowNo & IIf(strTag = "", "", " And Tag='" & strTag & "'")
                        If gclsPros.SecdInfoRec.EOF Then
                            Call gclsPros.SecdInfoRec.AddNew(Array("Sort", "���", "�ı�״̬", "ID", "ҳ��", "�ؼ���", "IndexEx", "��Ϣԭֵ", "����Ϣԭֵ", "Tag"), Array(Identity(lngSort), gclsPros.MainInfoRec!���, 0, IIf(lngID = 0, Null, lngID), gclsPros.MainInfoRec!ҳ��, gclsPros.MainInfoRec!�ؼ���, lngRowNo, strWholeInfo, strMainInfo, strTag))
                        End If
                    Else
                        'TagΪ�յ�������Ϣ�Թ���������������TagΪ����������Tag������ҪӦ���ڲ���������Ŀ����ʱ�洢"��,��"
                        If strInfoName = "������Ŀ" Then
                            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And Tag='" & Trim(strTag) & "'"
                        Else
                            gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And ����Ϣԭֵ=" & IIf(strMainInfo = "", "Null", "'" & strMainInfo & "'") & IIf(strTag <> "", " And Tag='" & Trim(strTag) & "'", "")
                        End If
                        If Not gclsPros.SecdInfoRec.EOF Then
                            If gclsPros.SecdInfoRec.RecordCount > 1 Then
                                gclsPros.SecdInfoRec.MoveFirst
                                gclsPros.SecdInfoRec.Filter = "ID=" & gclsPros.SecdInfoRec!ID
                            End If
                            Call gclsPros.SecdInfoRec.Update(Array("IndexEx", "��Ϣ��ֵ", "����Ϣ��ֵ", "Tag"), Array(lngRowNo, strWholeInfo, strMainInfo, IIf(strTag = "", gclsPros.SecdInfoRec!Tag, strTag)))
                        Else
                            Call gclsPros.SecdInfoRec.AddNew(Array("Sort", "���", "�ı�״̬", "ҳ��", "�ؼ���", "IndexEx", "��Ϣ��ֵ", "����Ϣ��ֵ", "Tag", "�ı�״̬"), Array(Identity(lngSort), gclsPros.MainInfoRec!���, 0, gclsPros.MainInfoRec!ҳ��, gclsPros.MainInfoRec!�ؼ���, lngRowNo, strWholeInfo, strMainInfo, strTag, CS_������))
                        End If
                    End If
            End Select
        End If
    Else
        '������Ϣ��ֱ�ӹ��ˣ��ڱ���ʱʹ��
        If Not grsDeliceryInfo Is Nothing Then
            grsDeliceryInfo.Filter = "����=0"
            For i = 1 To grsDeliceryInfo.RecordCount
                If grsDeliceryInfo!��Ϣֵ <> grsDeliceryInfo!��Ϣ��ֵ Then
                    grsDeliceryInfo.Update "��¼����", 1
                End If
                grsDeliceryInfo.Update
            Next
            grsDeliceryInfo.Filter = "��¼����=1": grsDeliceryInfo.Sort = "��Ϣ��"
            grsBabyInfo.Filter = "��¼����=1"
            grsBabyDiag.Filter = "��¼����=1"
        End If
        '���´μ���Ϣ��¼�����ж�����Ϣ��¼��
        gclsPros.SecdInfoRec.Filter = ""
        gclsPros.SecdInfoRec.Sort = "Sort"
        For i = 1 To gclsPros.SecdInfoRec.RecordCount
            lng״̬ = CS_δ�ı�
            If gclsPros.SecdInfoRec!��Ϣԭֵ & "" <> gclsPros.SecdInfoRec!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(gclsPros.SecdInfoRec!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(gclsPros.SecdInfoRec!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And gclsPros.SecdInfoRec!����Ϣԭֵ & "" <> gclsPros.SecdInfoRec!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            If lng��� <> gclsPros.SecdInfoRec!��� And lng״̬ <> CS_δ�ı� Then
                Call Rec.Update(gclsPros.MainInfoRec, "���=" & gclsPros.SecdInfoRec!���, "�Ƿ�ı�", 1)
                lng��� = gclsPros.SecdInfoRec!���
            End If
            gclsPros.SecdInfoRec.Update "�ı�״̬", lng״̬
            gclsPros.SecdInfoRec.MoveNext
        Next
        
        '��������Ϣ��¼��
        gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=0 And ExpState=" & ES_������չ
        gclsPros.MainInfoRec.Sort = "���"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            If gclsPros.MainInfoRec!��Ϣԭֵ & "" <> gclsPros.MainInfoRec!��Ϣ��ֵ & "" Then
                gclsPros.MainInfoRec.Update "�Ƿ�ı�", 1
            End If
            gclsPros.MainInfoRec.MoveNext
        Next

        gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
        If gclsPros.PatiType = PF_���� And gclsPros.FuncType = fҽ����ҳ Then
             gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF Or gclsPros.CurrentForm.UCPatiVitalSigns.GetSaveSQL(gclsPros.����ID, gclsPros.��ҳID) <> "" Or gclsPros.IsLastDiag
        ElseIf gclsPros.FuncType = f������ҳ Then
            gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF
            If Not grsDeliceryInfo Is Nothing Then
                gclsPros.InfosChange = gclsPros.InfosChange Or Not grsDeliceryInfo.EOF
            End If
            If Not grsBabyInfo Is Nothing Then
                gclsPros.InfosChange = gclsPros.InfosChange Or Not grsBabyInfo.EOF
                If Not grsBabyDiag Is Nothing Then
                    gclsPros.InfosChange = gclsPros.InfosChange Or Not grsBabyDiag.EOF
                End If
            End If
        Else
            gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF
        End If
        '�����Ϣ��������Ϣ,������Ϣ�����Ǵ�������Դ��ȡ�������ǣ��ҽ���ı䣬����Ҫȫ��񱣴�
        If gclsPros.InfosChange Or gclsPros.DiagSel Then
            strTmp = "��ҽ���;��ҽ���;����ҩ��;�������"
            arrTmp = Split(strTmp, ";")
            For i = LBound(arrTmp) To UBound(arrTmp)
                gclsPros.MainInfoRec.Filter = "��Ϣ��='" & arrTmp(i) & "'"
                If Not gclsPros.MainInfoRec.EOF Then
                    '�������޸Ļ����滻��δ�ı��������Դ������о���������
                    Call Rec.Update(gclsPros.SecdInfoRec, "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬>=0 And Tag<> '" & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "'", "�ı�״̬", CS_������)
                    'ɾ����������Դ������о���������
                    Call Rec.Update(gclsPros.SecdInfoRec, "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬<0 And Tag<> '" & IIf(gclsPros.FuncType = f������ҳ, 4, 3) & "'", "�ı�״̬", CS_δ�ı�)
                    If gclsPros.IsLastDiag And arrTmp(i) Like "*���" Then
                        '�������޸Ļ����滻��δ�ı���о���������
                        Call Rec.Update(gclsPros.SecdInfoRec, "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬>=0", "�ı�״̬", CS_������)
                         '�ϴιҺ���ϣ�ɾ���в��ô���
                        Call Rec.Update(gclsPros.SecdInfoRec, "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬<0", "�ı�״̬", CS_δ�ı�)
                    End If
                    gclsPros.SecdInfoRec.Filter = "���=" & gclsPros.MainInfoRec!��� & " And �ı�״̬<>0"
                    If Not gclsPros.SecdInfoRec.EOF Then
                        gclsPros.MainInfoRec.Update "�Ƿ�ı�", 1
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    Debug.Print "UpdateCacheRecInfo:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetAllerInput(ByVal LngRow As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'���ܣ��������ҩ�������
'������strTYTInput=̫Ԫͨ������ҩ�ӿڷ��ص��ַ���
'    Dim strSql As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String

    With gclsPros.CurrentForm.vsAller

        strAllerOld = .Cell(flexcpData, LngRow, AI_����ҩ��) & ";" & .TextMatrix(LngRow, AI_����Դ����)

        If gclsPros.UseTYT Then
            arrTmp = Split(strTYTInput, ";")

            If UBound(arrTmp) < 1 Then Exit Sub
            If strAllerOld <> strTYTInput Or Val(.RowData(LngRow) & "") <> 0 Then
                .TextMatrix(LngRow, AI_����ҩ��) = arrTmp(1)
                .TextMatrix(LngRow, AI_����Դ����) = arrTmp(0)
                .RowData(LngRow) = 0
            End If
        Else
            If gclsPros.FuncType <> f������ҳ Then
                If gclsPros.CurrentForm.optAller(PC_��ҩƷĿ¼����).Value Then
                    If Not rsInput Is Nothing Then
                        .RowData(LngRow) = CLng(rsInput!ID)
                        .TextMatrix(LngRow, AI_����ҩ��) = NVL(rsInput!����)
                    Else
                        .RowData(LngRow) = 0
                        .TextMatrix(LngRow, AI_����ҩ��) = .EditText
                    End If
    
                    strAllerNew = .TextMatrix(LngRow, AI_����ҩ��) & ";" & .TextMatrix(LngRow, AI_����Դ����)
    
                    If strAllerOld <> strAllerNew Or Val(.RowData(LngRow) & "") <> 0 Then
                        .TextMatrix(LngRow, AI_����Դ����) = ""
                    End If
                Else
                    If Not rsInput Is Nothing Then
                        .TextMatrix(LngRow, AI_����ҩ��) = rsInput!���� & ""
                        .TextMatrix(LngRow, AI_����Դ����) = rsInput!���� & ""
                        .RowData(LngRow) = 0
                    Else
                        .RowData(LngRow) = 0
                        .TextMatrix(LngRow, AI_����ҩ��) = .EditText
                    End If
                End If
            Else
                If Not rsInput Is Nothing Then
                    .RowData(LngRow) = CLng(rsInput!ID)
                    .TextMatrix(LngRow, AI_����ҩ��) = NVL(rsInput!����)
                Else
                    .RowData(LngRow) = 0
                    .TextMatrix(LngRow, AI_����ҩ��) = .EditText
                End If

                strAllerNew = .TextMatrix(LngRow, AI_����ҩ��) & ";" & .TextMatrix(LngRow, AI_����Դ����)

                If strAllerOld <> strAllerNew Or Val(.RowData(LngRow) & "") <> 0 Then
                    .TextMatrix(LngRow, AI_����Դ����) = ""
                End If
            End If
        End If

        .Cell(flexcpData, LngRow, AI_����ҩ��) = .TextMatrix(LngRow, AI_����ҩ��)
        .TextMatrix(LngRow, AI_ҩ��ID) = Val(.RowData(LngRow) & "")
'        If .Cell(flexcpData, LngRow, AI_����ʱ��) = "" Then
'            curDate = zlDatabase.Currentdate
'            .TextMatrix(LngRow, AI_����ʱ��) = Format(curDate, "yyyy-MM-dd")
'            .Cell(flexcpData, LngRow, AI_����ʱ��) = Format(curDate, "yyyy-MM-dd")
'        End If
        'ʼ�ձ���һ����
        If LngRow = .Rows - 1 Then
            .AddItem "", LngRow + 1
            Call ChangeVSFHeight(gclsPros.CurrentForm.vsAller, True, 0)
        End If
    End With
End Sub

Public Sub AllerEnterNextCell()
    Dim i As Long, j As Long

    With gclsPros.CurrentForm.vsAller
        If .Col = AI_����ʱ�� Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AI_����ҩ��
                 Call .Select(.Row, .Col)
                 '��������������Ļ���������ʱ����ֹ��������
'                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Public Sub zlVsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngOldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����иı�ʱ,������ص���ɫ
    '��Σ�CustomColor-�Զ�����ɫ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:22:38
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '�иı�ʱ
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Public Sub zlVsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Զ���ɫ
    '���ƣ����˺�
    '���ڣ�2010-03-23 10:52:23
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '����ؼ�
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '���ѡ����ɫ
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zlVsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef LngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim LngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long

    Err = 0: On Error GoTo Errhand:

    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
    If lng���� <> -1 Then
        LngCol = lng����
    Else
        LngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If LngCol = -1 Then LngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    LngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = LngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, LngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        LngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = LngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, LngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    LngRow = .Row
                                End If
                            End If
                            .Col = LngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub

Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal LngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:������
    '����:vsGrid-�����е�������
    '     lngRow-��ǰ��
    '     blnBefor-��lngrow֮���֮��.true:֮��,false-֮��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", LngRow
        Else
            .AddItem "", LngRow + 1
        End If
        Call ChangeVSFHeight(vsGrid, True)
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = LngRow
            Else
                .Row = LngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, ByRef LngRow As Long, ByRef LngCol As Long, ByRef intKeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '����:ֻ���������ֺͻس����˸�
    '����:
    '   objctl:Vsgrid8.0�ؼ�
    '   intKeyascii:
    '           Keyascii:8 (�˸�)
    '   Row-��ǰ��
    '   Col-��ǰ��
    '   TextType:(0-�ı�ʽ;1-����ʽ;2-���ʽ)
    '����:һ��KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:

    If TextType = m�ı�ʽ Then
        If intKeyAscii = Asc("'") Then
            intKeyAscii = 0
        End If
        Exit Sub
    End If

    If intKeyAscii < Asc("0") Or intKeyAscii > Asc("9") Then
        Select Case intKeyAscii
        Case vbKeyReturn       '�س�

        Case 8                 '�˸�

        Case Asc(".")
            If TextType = m���ʽ Or TextType = m�����ʽ Then
                If InStr(objCtl.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    intKeyAscii = 0
                End If
            Else
                intKeyAscii = 0
            End If
        Case Asc("-")          '����
            Dim iRow As Long
            Dim icol As Long
            If Trim(objCtl.EditText) = "" Then Exit Sub
            If TextType <> m�����ʽ Then intKeyAscii = 0: Exit Sub
            If objCtl.EditSelStart <> 0 Then intKeyAscii = 0: Exit Sub      '��겻���һλ,�������븺��
            If InStr(1, objCtl.EditText, "-") <> 0 Then   'ֻ�ܴ���һ������
                intKeyAscii = 0
            End If
        Case Else
            intKeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    intKeyAscii = 0
End Sub


Public Sub TSJCSetDiagInput(ByVal LngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������������Ŀ������
    With gclsPros.CurrentForm.vsTSJC
        If Not rsInput Is Nothing Then
            .TextMatrix(LngRow, 1) = NVL(rsInput!����)
        Else
            .TextMatrix(LngRow, 1) = .EditText
        End If
        .Cell(flexcpData, LngRow, 1) = .TextMatrix(LngRow, 1)
    End With
End Sub

Public Sub TSJCEnterNextCell()
    With gclsPros.CurrentForm.vsTSJC
        If .Row = .Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If .Row + 1 > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Public Function DiagCellEditable(ByRef vsDiagTmp As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim bln��ҽ As Boolean
    Dim blnJudge As Boolean
    Dim dtTmp As DiagType
    Dim lng��ԺRow As Long

    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        '�����в��ɱ༭
        If .ColHidden(LngCol) Then Exit Function
        '��ҳ�Ѿ�ǩ�������ѡ�����в������޸����
        If gclsPros.FuncType = f���ѡ�� And gclsPros.IsSigned And LngCol <> DI_���� Then Exit Function
        '���������������������������������(�����߼�����������
        If .TextMatrix(LngRow, DI_�������) = "" Then
            If Not (gclsPros.FuncType = f������ҳ And LngCol = DI_��ϱ��� Or LngCol = DI_Del Or LngCol = DI_�������) Then Exit Function
        ElseIf gclsPros.FuncType <> f������ҳ Then
            If LngCol <> DI_���� And LngCol <> DI_Del And LngCol <> DI_���� Then
                '����ҽ�����ɱ༭
                If gclsPros.FuncType = f���ѡ�� Then '��Ҫ�ų���ǰҽ�������뵥
                    If GetAdviceIDByDiag(.TextMatrix(LngRow, DI_ҽ��IDs), Val(.RowData(LngRow))) <> "" Then Exit Function
                Else
                    If .TextMatrix(LngRow, DI_ҽ��IDs) <> "" Then Exit Function
                End If
            End If
        End If
        Select Case LngCol
            Case DI_�������
                If gclsPros.FuncType <> f������ҳ Then
                    '�ٴ�·����ϲ������
                    If gclsPros.PathState = PS_ִ���� Or gclsPros.PathState = PS_�������� Then
                        If Not CheckMergePath(gclsPros.����ID, gclsPros.��ҳID, Val(.TextMatrix(LngRow, DI_��Ϸ���)), Val(.TextMatrix(LngRow, DI_����ID))) Then Exit Function
                    End If
                    '����·�����ϣ�������ϲ������
                    If gclsPros.PathDiag <> "" And gclsPros.PathState > PS_�����ϵ��� Then
                        If InStr("," & gclsPros.PathDiag & ",", "," & .TextMatrix(.Row, DI_��Ϸ���) & "|" & Val(.TextMatrix(.Row, DI_����ID)) & "|" & Val(.TextMatrix(.Row, DI_���ID)) & ",") > 0 Then
                            Exit Function
                        End If
                    End If
                    '������ɵĳ�Ժ��ϲ������
                    If gclsPros.PathState = PS_�������� And gclsPros.PathOutTime Then
                        If bln��ҽ Then
                            blnJudge = .TextMatrix(.Row, DI_�������) = "��Ժ���" And gclsPros.InPath <= DT_��Ժ���XY
                        Else
                            blnJudge = .TextMatrix(.Row, DI_�������) = "��Ժ���" And gclsPros.InPath >= DT_�������ZY
                        End If
                        If blnJudge Then Exit Function
                    End If
                End If
            Case DI_��ϱ���
                '������ҳ��ϱ����������໥����������������ϱ��루Ϊ�˱�֤��ҳ����ȡ������¼����Ͼ�����ϱ�������⣩
                '��������������룬Ϊ�������ʱ��鿴�������
                If gclsPros.FuncType <> f������ҳ Then
                    Exit Function
                End If
            Case DI_ICD����
                '������ҳ����������Ժ�������븽��,���������ڹ̶����룬ͬ���������븽��
                If Not bln��ҽ Then Exit Function
                If .TextMatrix(.Row, DI_�̶�����) = "1" Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_�������XY Or _
                    Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY Then
                    Exit Function
                End If
            Case DI_��Ժ���
                If bln��ҽ Then
                    '��Ժ��Ϻ�Ժ�ڸ�Ⱦ���������Ժ���(��Ϊ����Ժ�ڸ�Ⱦ�ڳ�Ժʱ�Ѿ���ת��������)
                    blnJudge = Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_Ժ�ڸ�Ⱦ Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_����֢
                Else
                    '�ǳ�Ժ���ʱ����������
                    blnJudge = Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���ZY
                End If
                If Not blnJudge Then Exit Function
                If gclsPros.FuncType = f������ҳ Then
                    If .TextMatrix(LngRow, DI_�Ƿ���) <> "1" Then Exit Function
                End If
            Case DI_��Ժ����
                '��Ժ����ֻ���ڳ�Ժ��Ϻ������������д,��ҽ�Ĳ���֢��Ժ�ڸ�ȾҲ������д
                If bln��ҽ Then
                    If Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_��Ժ���XY And Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_����֢ And Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_Ժ�ڸ�Ⱦ Then Exit Function
                Else
                    If Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_��Ժ���ZY Then Exit Function
                End If
            Case DI_�Ƿ�δ�� '��ҽδ��������
                blnJudge = Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_Ժ�ڸ�Ⱦ Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_����֢
                '��Ժ��Ϻ�Ժ�ڸ�Ⱦ���������Ƿ�δ��(��Ϊ����Ժ�ڸ�Ⱦ�ڳ�Ժʱ�Ѿ���ת��������)
                If Not blnJudge Then Exit Function
                '��Ժ���Ϊ"����"ʱ�ſ��������Ƿ�δ��
                If .TextMatrix(LngRow, DI_��Ժ���) <> "����" Then Exit Function
            Case DI_����
                '��Ժ��Ҫ��ϲ���������
                If bln��ҽ Then
                    blnJudge = .TextMatrix(LngRow, DI_�������) = "��Ժ���" And Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY
                Else
                    blnJudge = .TextMatrix(LngRow, DI_�������) = "��Ҫ���" And Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���ZY
                End If
                If blnJudge Then Exit Function
                'ͬ������һ�����Ϊ�գ�����������
                If LngRow <> .Rows - 1 Then
                    blnJudge = .TextMatrix(LngRow, DI_��Ϸ���) = .TextMatrix(LngRow + 1, DI_��Ϸ���) And .TextMatrix(LngRow, DI_�������) <> "" And .TextMatrix(LngRow + 1, DI_�������) = ""
                    If blnJudge Then Exit Function
                End If
        End Select

        '��Ժ��ϱ�����������(��δ����ʱ)
        dtTmp = IIf(bln��ҽ, DT_��Ժ���XY, DT_��Ժ���ZY)
        If .TextMatrix(LngRow, DI_�������) = "" And Val(.TextMatrix(LngRow, DI_��Ϸ���)) = dtTmp Then
            If .TextMatrix(LngRow - 1, DI_�������) = "" And Val(.TextMatrix(LngRow - 1, DI_��Ϸ���)) = dtTmp Then
                Exit Function
            End If
        End If

        If gclsPros.FuncType = f������ҳ And bln��ҽ Then
            lng��ԺRow = FindDiagRow(DT_��Ժ���XY)
            If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_�����ж��� Then
                If (.TextMatrix(lng��ԺRow, DI_��ϱ���) = "" And .TextMatrix(lng��ԺRow, DI_�������) = "") Or InStr("ST", Left(.TextMatrix(lng��ԺRow, DI_��ϱ���), 1)) = 0 Then

                ElseIf .TextMatrix(LngRow, DI_�̶�����) = "1" And InStr("VWXY", Left(.TextMatrix(LngRow, DI_ICD����), 1)) > 0 Then
                '�̶����벻�����޸�
                    Exit Function
                End If
            End If
        End If
        DiagCellEditable = True
    End With
End Function

Public Function GetMedInputSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str�Ա� As String, Optional ByVal strOtherInfo As String) As String
'���ܣ���ò�ѯ��ҳ�����ѯ��SQL
'������intType:��ȡ��SQL����,0-��ҽ��ϣ�1-��ҽ��ϣ�2-��������
'    strInput-��ѯ������str�Ա�--���˵��Ա�
'   strOtherInfo:��ҽ���-�����������ࣻ��ҽ���-�������
'���أ�strsql--��ѯ��ϵ�SQL
    Dim strSql As String

    If gclsPros.Sex Like "*��*" Then
        str�Ա� = "��"
    ElseIf gclsPros.Sex Like "*Ů*" Then
        str�Ա� = "Ů"
    End If

    Select Case intType
        Case 0, 1 '��ҽ���,��ҽ���
            If intType = 0 And gclsPros.DiagInputXY = 0 Or intType = 1 And gclsPros.DiagInputZY = 0 And strOtherInfo <> "Z" Then
            '���������:һ����Ͽ������ڶ������
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                strSql = "Select A.Id, A.Id ��ĿID, A.����, Null ���, Null ����, Null ����id, Null ��������, A.����, A.˵��, A.����, B.����, 0 ��Ч����, 0 ����," & vbNewLine & _
                                "              0 �Ƿ���, Max(D.����id) ����id, A.Id ���id" & vbNewLine & _
                                "       From �������Ŀ¼ A, ������ϱ��� B, ������϶��� D" & vbNewLine & _
                                " Where A.ID=B.���ID And A.ID=D.���ID(+) And A.���=" & IIf(intType = 0, 1, 2) & vbNewLine & _
                                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                " And B.����=[5] And (" & strSql & ")" & vbNewLine & _
                                "Group By A.Id, A.����, A.����, A.˵��, A.����,B.����"
                '��ȡ��϶�Ӧ�������븽��
                strSql = "Select distinct A.ID,A.��ĿID, A.����, B.���, B.����, Null ����id, Null ��������, A.����, A.˵��, Null ����,A.����, A.��Ч����, A.����, A.�Ƿ���," & vbNewLine & _
                                "       B.���� ��������, B.Id ����id, B.��� �������, A.���id," & vbNewLine & _
                                "                 Decode(a.����, [6], 1, Decode(A.����,[6],1,decode(A.����,[6],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                                "                 Decode(Substr(A.����, 1, Length([6])), [6], 1, Decode(Substr(A.����, 1, Length([6])),[6],1,decode(Substr(a.����, 1, Length([6])),[6],1,NULL))) As ����3ID" & _
                                " From (" & strSql & ") A, ��������Ŀ¼ B, ������Ͽ��� C, ������Ͽ��� D" & vbNewLine & _
                                " Where A.����id = B.Id(+)" & vbNewLine & _
                                " And c.���id(+) = a.Id And d.���id(+) = a.Id And c.����id(+)=[8]  And d.��Աid(+) = [7]" & _
                                " Order By ����1ID, ����2ID, ����3ID, A.����"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "A.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSql = "A.���� Like [1] Or A.���� Like [2] Or " & IIf(gclsPros.BriefCode = 0, "A.����", "A.�����") & " Like [2]"
                End If
                If gclsPros.FuncType = f������ҳ Then
                    strSql = _
                        "Select A.Id, A.Id ��ĿID,A.����, A.���, A.����,Null ����ID, Null ��������, A.����, A.˵��, Null ����,A.����id, " & IIf(gclsPros.BriefCode = 0, "A.����", "A.�����") & " as ����,  A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������, Null ���id" & vbNewLine & _
                        "From ��������Ŀ¼ A, ����������� C" & vbNewLine & _
                        "Where A.����id = C.Id(+) And Instr([3],A.���)>0 And (" & strSql & ")" & _
                        IIf(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is NULL)", "") & _
                        " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by A.����"
                Else
                    strSql = _
                        "Select A.Id,A.Id ��ĿID, A.����, A.���, A.����,Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, " & IIf(gclsPros.BriefCode = 0, "A.����", "A.�����") & " as ����,  A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������," & vbNewLine & _
                        "       Max(B.���id) ���id" & vbNewLine & _
                        "From ��������Ŀ¼ A, ������϶��� B, ����������� C " & vbNewLine & _
                        "Where A.Id = B.����id(+) And A.����id = C.Id(+)  And" & vbNewLine & _
                        " Instr([3],A.���)>0 And (" & strSql & ")" & _
                        IIf(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is NULL)", "") & _
                        " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        "Group By A.Id, A.����, A.���, A.����, A.����, A.˵��, A.����id, " & IIf(gclsPros.BriefCode = 0, "A.����", "A.�����") & ", A.��Ч����, A.����, A.���,C.�Ƿ���"
                End If
                strSql = "Select distinct A.Id,A.��ĿID, A.����, A.���, A.����,A.����ID, A.��������, A.����, A.˵��, A.����, A.����id, A.����,  A.��Ч����, A.����, A.�Ƿ���,A.��������, A.����id,A.�������,A.���id, " & _
                        " Decode(a.����, [6], 1, Decode(A.����,[6],1,decode(a.����,[6],1,NULL))) As ����1ID," & vbNewLine & _
                "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                "                Decode(Substr(a.����, 1, Length([6])), [6], 1, Decode(Substr(A.����, 1, Length([6])),[6],1,decode(Substr(a.����, 1, Length([6])),[6],1,NULL))) As ����3ID" & vbNewLine & _
                        " From (" & strSql & ") A, ����������� C, ����������� D " & _
                        " Where  c.����id(+) = a.Id And d.����id(+) = a.Id And c.����id(+)=[8]  And d.��Աid(+) = [7] " & _
                        " Order By" & IIf(strOtherInfo = "'M,D'", " ������� desc , ", "") & " ����1ID, ����2ID, ����3ID, A.����"
            End If
        Case 2 '��������
            If gclsPros.OPSInput = 0 And gclsPros.FuncType <> f������ҳ Then
                '��������Ŀ����
                strSql = "Select distinct A.ID,A.����,A.����,A.�������� as ��ģ" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.���='F' And A.������� IN(2,3) And A.ID=B.������ĿID" & _
                    IIf(str�Ա� <> "", " And Nvl(A.�����Ա�,0) IN(0,[4])", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or B.���� Like [2] Or B.���� Like [2])" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
            Else
                '��ICD9-CM3����
                strSql = " Select distinct ID,����,����,����,����,˵��" & _
                    " From ��������Ŀ¼ Where ���='S'" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                    " Order by ����"
            End If
    End Select
    GetMedInputSQL = strSql
End Function

Public Function OPSCellEditable(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim vsTemp As VSFlexGrid
    Set vsTemp = gclsPros.CurrentForm.vsOPS

    With vsTemp
        If .ColHidden(LngCol) Then Exit Function

        '������������������,��������
        If Not IsDate(.TextMatrix(LngRow, PI_��������)) Then
            If LngCol > PI_�������� Then Exit Function
        End If
        If .TextMatrix(LngRow, PI_��������) = "" Then
            If LngCol > PI_�������� Then Exit Function
        End If

        '��������������ҽʦ
        If .TextMatrix(LngRow, PI_����ҽʦ) = "" Then
            If LngCol = PI_����1 Or LngCol = PI_����2 Then Exit Function
        End If

        '�����������1����
        If .TextMatrix(LngRow, PI_����1) = "" Then
            If LngCol = PI_����2 Then Exit Function
        End If

        '������������������
        If Trim(.TextMatrix(LngRow, PI_��������)) = "" Then
            If LngCol = PI_����ҽʦ Then Exit Function
        End If
        If gclsPros.FuncType <> f������ҳ Then
            '�������Ʋ�������
            If LngCol = PI_�������� And gclsPros.CurrentForm.chkParaOPSInfo(PC_δ�ҵ�ʱ����¼��).Value = 0 Then Exit Function
        Else
            If LngCol = PI_�������� And Not gclsPros.CNIndent Then Exit Function
        End If

        '��ȡ�����������������
        If LngCol = PI_�������� Then
            If gclsPros.Module = pסԺҽ��վ Then
                If .Cell(flexcpData, LngRow, LngCol) = 1 And InStr(GetInsidePrivs(pסԺҽ��վ), "�޸������ȼ�") = 0 Then Exit Function
            Else
                If .Cell(flexcpData, LngRow, LngCol) = 1 Then Exit Function
            End If
        End If
        
    End With
    OPSCellEditable = True
End Function

Public Function CheckIsDate(ByVal strKEY As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:(20070101��2007-01-01)����(01-01��0101)����(01<01-31>)
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKEY) = 4 And InStr(1, strKEY, "-") = 0 Then
        '0101,��Ҫ��ǰ�����
        strKEY = Year(Now) & strKEY
    ElseIf Len(Replace(strKEY, "-", "")) = 4 And InStr(1, strKEY, "-") > 0 Then
        '01-01��ʽ,��Ҫ����
        strKEY = Year(Now) & Replace(strKEY, "-", "")
    ElseIf Len(strKEY) <= 2 And IsNumeric(strKEY) Then
        'ָ����
        strKEY = Format(Now, "YYYYMM") & IIf(Len(strKEY) = 2, strKEY, "0" & strKEY)
    End If
    If Len(strKEY) = 8 And InStr(1, strKEY, "-") = 0 Then
        strKEY = TranNumToDate(strKEY)
        If strKEY = "" Then
            MsgBox strTittle & "����Ϊ������,���飡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Not IsDate(strKEY) Then
        MsgBox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡", vbInformation, gstrSysName
        Exit Function
    End If
    CheckIsDate = strKEY
End Function

Public Function Check������Ч��(ByVal strDate As String, ByVal strTittle As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ڵ���Ч��
    '���:strDate-��ǰ����
    '     strTittle-����:��:�����ڵڼ���
    '����:
    '����:��Ч��strDate="",����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 17:03:30
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strCurDate As String

    If strDate = "" Then Check������Ч�� = True: Exit Function
    '��������Ƿ�Ϸ�
    If IsDate(strDate) = False Or IsNumeric(strDate) Then
        MsgBox strTittle & "����һ����Ч�����ڷ�Χ,����!", vbInformation, gstrSysName
        Exit Function
    End If

    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    If CDate(strDate) > CDate(strCurDate) Then
        MsgBox strTittle & "�ȵ�ǰ���ڻ�Ҫ��,����!", vbInformation, gstrSysName
        Exit Function
    End If

    If CDate(strDate) < CDate(gclsPros.InTime) Then
        MsgBox strTittle & "����Ժ���ڻ�ҪС,����!", vbInformation, gstrSysName
        Exit Function
    End If

    If gclsPros.OutTime <> "" Then
        If CDate(gclsPros.OutTime) < CDate(strDate) Then
            MsgBox strTittle & "�ȳ�Ժ���ڻ�Ҫ��,����!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check������Ч�� = True
End Function

Public Function DblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     blnNegative     �Ƿ���и������
    '     blnZero         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------

    Dim dblValue As Double

    If blnZero = True Then
        If strInput = "" Then
            MsgBox str��Ŀ & "δ���룬����!", vbInformation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then DblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If

    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If

    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If

    If blnZero = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    DblIsValid = True
End Function

Public Function CheckInPutIsDate(ByVal vsObj As Object, LngRow As Long, LngCol As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:���������������Ƿ�Ϸ�
    '����:lngRow -��,lngCol -��
    '����:���ںϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/05/21
    '---------------------------------------------------------------------------------------------------------
    Dim strKEY As String
    Dim str����ʱ�� As String, str�˳�ʱ�� As String

    strKEY = Trim(vsObj.EditText)
    strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
    strKEY = Replace(strKEY, Chr(10), "")
    strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
    If strKEY <> "" Then
        If Not IsDate(strKEY) Then
            MsgBox vsObj.TextMatrix(0, LngCol) & "����Ϊ������,���������룡", vbInformation + vbDefaultButton1, gstrSysName
             vsObj.EditSelStart = 0
             vsObj.EditSelLength = 1000
            Exit Function
        End If

        Select Case LngCol
           Case UI_����ʱ��
               str����ʱ�� = strKEY
               str�˳�ʱ�� = Trim(vsObj.TextMatrix(LngRow, UI_�˳�ʱ��))
               If str�˳�ʱ�� <> "" And str����ʱ�� >= str�˳�ʱ�� Then
                   MsgBox "ע:" & vbCrLf & "  ����ʱ��������˳�ʱ��,���飡", vbInformation + vbDefaultButton1, gstrSysName
                   Exit Function
               End If
           Case UI_�˳�ʱ��
               str����ʱ�� = Trim(vsObj.TextMatrix(LngRow, UI_����ʱ��))
               str�˳�ʱ�� = strKEY

               If str����ʱ�� <> "" And CDate(str����ʱ��) >= CDate(str�˳�ʱ��) Then
                   MsgBox "ע:" & vbCrLf & "  �˳�ʱ��С���˽���ʱ��,���飡", vbInformation + vbDefaultButton1, gstrSysName
                   Exit Function
               End If
        End Select
    End If
    CheckInPutIsDate = True
End Function

Public Sub ChangePage(Optional ByVal blnForWord As Boolean = True, Optional ByVal lngPage As Long = -1, Optional ByRef objTmp As Object, Optional blnLocation As Boolean = True)
'���ܣ�ѡ���λ����һҳ
'������blnForWord=�Ƿ���ǰ��ҳ�������һҳʱ����λ����һҳ��false=���ҳ
'      lngPage=ָ��ҳ��������-1ʱ�������з�ҳ��ֱ�Ӷ�λҳ��
    Dim lngCurPage As Long, i As Long
    Dim lngHeight As Long
    Dim lngMin As Long, lngMax As Long
    Dim blnCur As Boolean
  
    lngMin = -1
    With gclsPros.CurrentForm
        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                If lngMin = -1 Then lngMin = i
                If i > lngMax Then lngMax = i
                If Not blnCur Then
                    lngHeight = lngHeight + .PicPage(i).Height
                    If lngHeight > 500 + Abs(.picMain.Top) Then
                        lngCurPage = i
                        blnCur = True
                    End If
                End If
            End If
        Next
        If lngPage = -1 Then
            If blnForWord Then
                For i = .PicPage.LBound To .PicPage.UBound
                    If .PicPage(i).Tag = "true" Then
                        If i > lngCurPage Then
                            lngPage = i
                            Exit For
                        End If
                    End If
                Next
                If lngPage = -1 Then lngPage = lngMax
            Else
                For i = .PicPage.UBound To .PicPage.LBound Step -1
                    If .PicPage(i).Tag = "true" Then
                        If i < lngCurPage Then
                            lngPage = i
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        
        If lngPage < lngMin Then
            lngPage = lngMin
        ElseIf lngPage > lngMax Then
            lngPage = lngMax
        End If
       
        lngHeight = 0
        For i = .PicPage.LBound To lngPage - 1
            If .PicPage(i).Tag = "true" Then
                lngHeight = lngHeight + .PicPage(i).Height
            End If
        Next
    
         i = Abs((-500 - lngHeight - .PicPage(0).ScaleTop) / ((.picMain.Height + 1100 - .ScaleHeight)) * 1000)
        .vsbMain.Value = IIf(i > 1000, 1000, i)
        If Not objTmp Is Nothing Then
             zlControl.ControlSetFocus objTmp
             Exit Sub
        End If
        
        If blnLocation Then                     '�Ƿ�λ���ؼ�����������ʱ�򲻶�λ
            Select Case lngPage
                Case PIC_סԺ��ҳ
                    If gclsPros.FuncType = fҽ����ҳ Then
                        zlControl.ControlSetFocus .cboBaseInfo(BCC_���ʽ)
                    ElseIf gclsPros.FuncType = f������ҳ Then
                        If Not .txtSpecificInfo(SLC_סԺ��).Locked Then
                            zlControl.ControlSetFocus .txtSpecificInfo(SLC_סԺ��)
                        Else
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_���ʽ)
                        End If
                    End If
                Case PIC_������Ϣ
                        If gclsPros.FuncType = fҽ����ҳ Then
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_����)
                        ElseIf gclsPros.FuncType = f������ҳ Then
                            If .txtInfo(GC_����).Locked Then
                                zlControl.ControlSetFocus .cboBaseInfo(BCC_����)
                            Else
                                zlControl.ControlSetFocus .txtInfo(GC_����)
                            End If
                        End If
                Case PIC_��ҽ���
                    Call LocateVSFRowCol(.vsDiagXY, 1, .vsDiagXY.Rows - 1, DI_��ϱ���, DI_Del, 1, DI_�������)
                    zlControl.ControlSetFocus .vsDiagXY
                Case PIC_��ҽ������
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_��Ժ���)
                Case PIC_��ҽ���
                    Call LocateVSFRowCol(.vsDiagZY, 1, .vsDiagZY.Rows - 1, DI_��ϱ���, DI_Del, 1, DI_�������)
                    zlControl.ControlSetFocus .vsDiagZY
                Case PIC_��ҽ������
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_�������ԺZY)
                Case PIC_ҩ�����
                    Call LocateVSFRowCol(.vsAller, 1, .vsAller.Rows - 1, AI_����ҩ��, AI_����ʱ��, 1, AI_����ҩ��)
                    zlControl.ControlSetFocus .vsAller
                Case PIC_��Ѫ��Ϣ
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_Ѫ��)
                Case PIC_ǩ����Ϣ
                    zlControl.ControlSetFocus .cboManInfo(MC_������)
                Case PIC_������¼
                    Call LocateVSFRowCol(.vsOPS, 1, .vsOPS.Rows - 1, PI_��������, PI_�п�����, 1, PI_��������)
                    zlControl.ControlSetFocus .vsOPS
                Case PIC_סԺ����
                    zlControl.ControlSetFocus .chkFeeEdit
                Case PIC_סԺ���
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_��������)
                Case PIC_������Ϣ
                    zlControl.ControlSetFocus .vsChemoth
                Case PIC_���Ƽ�¼
                    zlControl.ControlSetFocus .vsRadioth
                Case PIC_������
                    zlControl.ControlSetFocus .vsSpirit
                Case PIC_����ҩ��
                    zlControl.ControlSetFocus .vsKSS
                Case PIC_��֢�໤
                    zlControl.ControlSetFocus .vsFlxAddICU
                Case PIC_��������
                    zlControl.ControlSetFocus .vsfMain
                Case PIC_��ҳ1
                    If gclsPros.MedPageSandard = ST_��������׼ Then
                         zlControl.ControlSetFocus .lstInfection
                    ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                        zlControl.ControlSetFocus .vsInfect
                    ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                        zlControl.ControlSetFocus .txtInfo(GC_��֢�໤������)
                    ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                        If .optInput(OP_ICU��).Value Then
                            zlControl.ControlSetFocus .optInput(OP_ICU��)
                        Else
                            zlControl.ControlSetFocus .optInput(OP_ICU��)
                        End If
                    End If
                Case PIC_��ҳ2
                    If gIntPic + 1 <> PIC_��ҳ2 Then
                        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                            zlControl.ControlSetFocus .chkInfo(CHK_����·��)
                        End If
                    Else
                        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                            zlControl.ControlSetFocus gfrmMecCol(lngPage - gIntPic)
                        End If
                    End If
                Case Else
                    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                            zlControl.ControlSetFocus gfrmMecCol(lngPage - gIntPic)
                    End If
            End Select
        End If
    End With
End Sub

Public Sub PrintInMedRec(ByVal mopType As MedRec_Operate, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, _
                        Optional ByVal intPage As Integer, Optional ByRef objReport As Object, Optional ByRef objForm As Object)
'���ܣ���ҳ��ӡ��Ԥ��
'������intType=2����ӡ����=1��Ԥ����0=����
'     lng����ID-���˿���
'     intPage=1-4��ӡ��ҳ������ʽ��=5��ӡ����+��ҳ1��=6��ӡ����+��ҳ2
    Dim strName As String
    Dim lngPage As Long
    Dim objReportTmp As clsReport
    Dim objFormTmp As Object
    Dim bln��ҽ As Boolean

    If lng����ID <> 0 Then
        If gobjReport Is Nothing Then Set gobjReport = New clsReport
        Set objReportTmp = IIf(objReport Is Nothing, gobjReport, objReport)
        Set objFormTmp = IIf(objForm Is Nothing, gclsPros.CurrentForm, objForm)
        bln��ҽ = sys.DeptHaveProperty(lng����ID, "��ҽ��")
        '����ϵͳ��ӡ����
        If gclsPros.SysNo \ 100 = 3 Then
            strName = "ZL3_BILL_200"
            intPage = 0
            mopType = MOP_��ӡ
        Else
            Select Case gclsPros.MedPageSandard
                Case ST_��������׼ '��������׼
                    If bln��ҽ Then
                        strName = "ZL1_INSIDE_1261_4"
                    Else
                        strName = "ZL1_INSIDE_1261_1"
                    End If
                Case ST_�Ĵ�ʡ��׼    '�Ĵ�ʡ��׼
                    If bln��ҽ Then
                        strName = "ZL1_INSIDE_1261_6"
                    Else
                        strName = "ZL1_INSIDE_1261_5"
                    End If
                Case ST_����ʡ��׼    '����ʡ��׼
                    If bln��ҽ Then
                        strName = "ZL1_INSIDE_1261_8"
                    Else
                        strName = "ZL1_INSIDE_1261_7"
                    End If
                Case ST_����ʡ��׼    '����ʡ��׼
                    If bln��ҽ Then
                        strName = "ZL1_INSIDE_1261_10"
                    Else
                        strName = "ZL1_INSIDE_1261_9"
                    End If
            End Select

            If GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.DBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 0) = 0 And intPage = 0 Then
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.DBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 1)
            End If
        End If
        
        If mopType = MOP_���� Then
            Call ReportPrintSet(gcnOracle, gclsPros.SysNo, strName, objFormTmp)
        Else
            If intPage = 5 Then
                lngPage = 1
            ElseIf intPage = 6 Then
                lngPage = 2
            Else
                lngPage = intPage
            End If
            Call objReportTmp.ReportOpen(gcnOracle, gclsPros.SysNo, strName, objFormTmp, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), mopType)
            If intPage > 4 Then
                Call objReportTmp.ReportOpen(gcnOracle, gclsPros.SysNo, strName, objFormTmp, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), mopType)
            End If
        End If
    End If
End Sub

Public Sub SetKSSSerial()
'���ܣ����ÿ���ҩ���������
    Dim i As Long

    With gclsPros.CurrentForm.vsKSS
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, KI_���) = i
        Next
    End With
End Sub

Public Sub SetPatiAddress(ByVal lngIndex As Long, ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal blnDefault As Boolean)
'���ܣ�����ĳ����ַ��ؿؼ���ֵ
'����:lngIndex=��ַ��ؿؼ�Index
'     strInfoName=��Ϣ��
'     strInfoVale��Ϣֵ
'     blnDefault=�Ƿ�������Ĭ��ֵ
    Dim rsTmp As ADODB.Recordset
    Dim blnHavePadr As Boolean '�Ƿ��в��˵�ַ�ؼ�
    Dim strTmp As String

    On Error GoTo errH
    With gclsPros.CurrentForm
        On Error Resume Next
        Err.Clear: strTmp = .padrInfo(lngIndex).Value
        blnHavePadr = Err.Number = 0: Err.Clear
        On Error GoTo errH
        If gclsPros.IsStructAdress And blnHavePadr Then
            Set rsTmp = GetStrucAddress(gclsPros.����ID, gclsPros.��ҳID, strInfoName)
            If rsTmp.RecordCount > 0 Then
                Call .padrInfo(lngIndex).LoadStructAdress(rsTmp!ʡ & "", rsTmp!�� & "", rsTmp!�� & "", rsTmp!���� & "", rsTmp!���� & "")
                If blnDefault Then .padrInfo(lngIndex).Tag = rsTmp!ʡ & "" & rsTmp!�� & "" & rsTmp!�� & "" & rsTmp!���� & "" & rsTmp!����
            Else
                .padrInfo(lngIndex).Value = strInfoValue
                If blnDefault Then .padrInfo(lngIndex).Tag = strInfoValue
            End If
        Else
            .txtAdressInfo(lngIndex).Text = strInfoValue
            If blnDefault Then .txtAdressInfo(lngIndex).Tag = strInfoValue
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ValidateAge(ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox, Optional ByVal bytIndex As Byte = 0) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
'61454:������,2013-05-14,��Ӷ�Ӥ�׶������У��
'bytIndex 0 �������� 1 Ӥ�׶�����
    If Not IsNumeric(txt����.Text) Then ValidateAge = True: Exit Function

    If bytIndex = 0 Then
        Select Case cbo���䵥λ.Text
            Case "��"
                If Val(txt����.Text) > 200 Then
                    MsgBox "����ֵ�������������200�꣬���������Ƿ���ȷ��", vbInformation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "��"
                If Val(txt����.Text) > 2400 Then
                    MsgBox "����ֵ�������������2400�£����������Ƿ���ȷ��", vbInformation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "��"
                If Val(txt����.Text) > 73000 Then
                    MsgBox "����ֵ�����������73000�죬���������Ƿ���ȷ��", vbInformation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "Сʱ" '���ܴ���30�켴720Сʱ
                If Val(txt����.Text) > 720 Then
                    MsgBox "����ֵ�������������720Сʱ����ʹ�ú��ʵ����䵥λ��", vbInformation, gstrSysName
                    If cbo���䵥λ.Enabled And cbo���䵥λ.Visible Then cbo���䵥λ.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "����" '���ܴ���24Сʱ��1440����
                If Val(txt����.Text) > 1440 Then
                    MsgBox "����ֵ�������������1440���ӣ���ʹ�ú��ʵ����䵥λ��", vbInformation, gstrSysName
                    If cbo���䵥λ.Enabled And cbo���䵥λ.Visible Then cbo���䵥λ.SetFocus
                    ValidateAge = False: Exit Function
                End If
        End Select
    Else
        Select Case cbo���䵥λ.Text
            Case "��"
                If Val(txt����.Text) > 12 Then
                    MsgBox "Ӥ������ֵ�������������12�£����������Ƿ���ȷ��", vbInformation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "��"
                If Val(txt����.Text) > 365 Then
                    MsgBox "Ӥ������ֵ�������������365�죬���������Ƿ���ȷ��", vbInformation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "Сʱ"
                If Val(txt����.Text) > 720 Then
                    MsgBox "Ӥ������ֵ�������������720Сʱ����ʹ�ú��ʵ����䵥λ��", vbInformation, gstrSysName
                    If cbo���䵥λ.Enabled And cbo���䵥λ.Visible Then cbo���䵥λ.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "����"
                If Val(txt����.Text) > 1440 Then
                    MsgBox "Ӥ������ֵ�������������1440���ӣ���ʹ�ú��ʵ����䵥λ��", vbInformation, gstrSysName
                    If cbo���䵥λ.Enabled And cbo���䵥λ.Visible Then cbo���䵥λ.SetFocus
                    ValidateAge = False: Exit Function
                End If
        End Select
    End If
    ValidateAge = True
End Function

Public Function GetKSSUseStage(ByVal DateUseBegin As Date, ByVal DateUseEnd As Date, ByVal DateSs As Date) As String
'���ܣ���ÿ�����ʹ�ý׶�
'������DateUseBegin ʹ��ʱ��,DateUseEnd -����ʱ��  DateSs-����ʱ��,strTime ��һ�ε�ʹ�ý׶�
    Dim strTimeTmp As String

    '���û���������򷵻ؿ�
    If DateSs <> CDate(0) Then
        If DateUseBegin < DateSs And DateUseEnd < DateSs Then
            strTimeTmp = "��ǰ"
        ElseIf DateUseBegin > DateSs And DateUseEnd > DateSs Then
            strTimeTmp = "����"
        ElseIf DateUseBegin = DateSs And DateUseEnd = DateSs Then
            strTimeTmp = "����"
        End If
        If strTimeTmp = "" Then strTimeTmp = "Χ������"
    End If
    GetKSSUseStage = strTimeTmp
End Function

Public Function GetKSSUseDay(ByVal AdviceID As Long, ByVal lngҩƷID As Long, ByVal strִ��ʱ�䷽�� As String, ByVal Date��ʼִ��ʱ�� As Date, _
            ByVal Date����ʱ�� As Date, ByVal lngƵ�ʴ��� As Long, ByVal lngƵ�ʼ�� As Long, ByVal str�����λ As String, ByVal str��ҩĿ�� As String, _
            ByRef rsTime As ADODB.Recordset) As Long
'���ܣ���ȡ�����ص�ʹ������
    Dim blnNew As Boolean
    Dim strPause As String
    Dim j As Long
    Dim StrDecTime As String, arrDecTime As Variant
    Dim DateStart As String
    Dim strTmp As String

    '��¼��ûʵ�����������¶���
    blnNew = rsTime Is Nothing
    If Not blnNew Then
        blnNew = rsTime.Fields.Count <> 3
        If Not blnNew Then blnNew = rsTime.Fields(0).Name <> "�շ�ʱ��" Or rsTime.Fields(1).Name <> "ҩƷID" Or rsTime.Fields(2).Name <> "��ҩĿ��"
    End If

    If blnNew Then
        Set rsTime = New ADODB.Recordset
        rsTime.Fields.Append "�շ�ʱ��", adVarChar, 10
        rsTime.Fields.Append "ҩƷID", adBigInt
        rsTime.Fields.Append "��ҩĿ��", adVarChar, 100
        rsTime.CursorLocation = adUseClient
        rsTime.LockType = adLockOptimistic
        rsTime.CursorType = adOpenStatic
        rsTime.Open
    End If

    strPause = GetAdvicePause(AdviceID)

    If strִ��ʱ�䷽�� <> "" Then
        StrDecTime = Calc���ڷֽ�ʱ��(Date��ʼִ��ʱ��, Date����ʱ��, strPause, strִ��ʱ�䷽��, lngƵ�ʴ���, lngƵ�ʼ��, str�����λ)
        arrDecTime = Split(StrDecTime, ",")
        For j = 0 To UBound(arrDecTime)
            strTmp = Format(arrDecTime(j), "yyyy-MM-dd")
            rsTime.Filter = "�շ�ʱ��='" & strTmp & "' And " & "ҩƷid=" & lngҩƷID & " And ��ҩĿ��='" & str��ҩĿ�� & "'"
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!�շ�ʱ�� = strTmp
                rsTime!ҩƷID = lngҩƷID
                rsTime!��ҩĿ�� = str��ҩĿ��
                rsTime.Update
            End If
        Next
    Else
        DateStart = CDate(Format(Date��ʼִ��ʱ�� & "", "yyyy-MM-dd"))
        Do While DateStart <= CDate(Format(Date����ʱ�� & "", "yyyy-MM-dd"))
            rsTime.Filter = "�շ�ʱ��='" & Format(CStr(DateStart), "yyyy-MM-dd") & "' And " & "ҩƷid=" & lngҩƷID & " And ��ҩĿ��='" & str��ҩĿ�� & "'"
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!�շ�ʱ�� = Format(CStr(DateStart), "yyyy-MM-dd")
                rsTime!ҩƷID = lngҩƷID
                rsTime!��ҩĿ�� = str��ҩĿ��
                rsTime.Update
            End If
            DateStart = CDate(DateStart) + 1
        Loop
    End If
    rsTime.Filter = "ҩƷid=" & lngҩƷID & " And ��ҩĿ��='" & str��ҩĿ�� & "'"
    GetKSSUseDay = rsTime.RecordCount
End Function

Public Function ClearPageContent() As Boolean
'���ܣ������������
    Dim ctlTmp As Control
    Dim i As Long, j As Long
    Dim rsTemp As New ADODB.Recordset
    '�ؼ��������壬�������Բ鿴
    Dim vsTmp As VSFlexGrid, txtTmp As TextBox, paTmp As PatiAddress
    Dim chkTmp As CheckBox, lstTmp As ListBox, cboTmp As ComboBox
    Dim lvwTmp As ListView, mskTmp As MaskEdBox, optTmp As OptionButton
    Dim vsbTmp As VScrollBar, hsbTmp As HScrollBar
    Dim arrTmp As Variant
    On Error GoTo errH
    If gclsPros.FuncType = f���Ӳ��� Then
        gblnCheck = True
        For Each ctlTmp In gclsPros.CurrentForm.Controls
            Select Case TypeName(ctlTmp)
                Case "TextBox" 'Լ120-140��
                    Set txtTmp = ctlTmp
                    If txtTmp.Index = GCA_��������ʬ�� Then
                        txtTmp.Tag = ""
                        txtTmp.Text = ""
                    Else
                        txtTmp.Text = txtTmp.Tag '������Ĭ��ֵ
                    End If
                Case "CheckBox" '��ѡ�ؼ���Լ��15-30��֮��
                    Set chkTmp = ctlTmp
                    chkTmp.Value = 0
                Case "VSFlexGrid" '������
                    Set vsTmp = ctlTmp
                    vsTmp.Clear
                Case "ListBox" '��3��
                    Set lstTmp = ctlTmp
                    lstTmp.Clear
            End Select
        Next
        gblnCheck = False
    Else
        For Each ctlTmp In gclsPros.CurrentForm.Controls
            'case������е��Ⱥ�˳�򣺽����ķ�����ǰ�棬�����ؼ����٣��ź���
            Select Case TypeName(ctlTmp)
                Case "Label", "Frame"
                    'lbl����������������ռλ�ã���Ϊlbl�ؼ��Ƚ϶࣬���Էŵ�һλ
                Case "TextBox" 'Լ50-60��
                    Set txtTmp = ctlTmp
                    txtTmp.Text = ""
                    '�ָ�Ĭ��ֵ
                    If txtTmp.Name = "txtAdressInfo" Then
                        txtTmp.Text = txtTmp.Tag
                    Else
                         txtTmp.Tag = ""
                    End If
                Case "ComboBox" 'Լ40-50��
                    Set cboTmp = ctlTmp
                    If cboTmp.Style = 0 Then
                        cboTmp.Text = "" '��������������б������������
                        cboTmp.Tag = ""
                    End If
                    If cboTmp.Tag <> "" Then '�ָ�Ĭ��ֵ
                        cboTmp.ListIndex = Val(cboTmp.Tag)
                    Else
                        cboTmp.ListIndex = -1
                    End If
                    '����ֹ���ӵ���Ա
                    If gclsPros.FuncType = f������ҳ And cboTmp.Name = "cboManInfo" And cboTmp.ListCount > 0 Then
                        For i = cboTmp.ListCount - 1 To 0 Step -1
                            If cboTmp.ItemData(i) = -999 Then
                                cboTmp.RemoveItem i
                            Else '��Ϊ�ǴӺ�����ӣ���˲���-999���˳�ѭ��
                                Exit For
                            End If
                        Next
                    End If
                Case "CheckBox" '��ѡ�ؼ���Լ��10-20��֮��
                    Set chkTmp = ctlTmp
                    '�ָ�Ĭ��ֵ
                    chkTmp.Value = 0
'                    If chkTmp.Index = CHK_�Ƿ�ȷ�� Then
'                        chkTmp.Value = 1
'                    Else
'                        chkTmp.Value = 0
'                    End If
                Case "VSFlexGrid" '������
                    Set vsTmp = ctlTmp
                    '�̶��в�����0��ֻ����ϱ��ת�Ƽ�¼�������飬��Щֻ��Ҫ�����Ԫ�����ݼ���
                    If vsTmp.FixedCols <> 0 Then
                        '�������������е�����
                        'ɾ��������Ϊ�յ��У����������Ϊ�յ��еĵ�Ԫ������
                        If vsTmp.Name = "vsDiagXY" Or vsTmp.Name = "vsDiagZY" Then
                            '�������
                            vsTmp.Cell(flexcpData, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = Empty
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, DI_��Ϸ��� - 1) = ""
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, DI_��Ϸ��� + 1, vsTmp.Rows - 1, vsTmp.Cols - 1) = ""
                            i = vsTmp.FixedRows: j = vsTmp.Rows - 1
                            Do While i <= j
                               If vsTmp.TextMatrix(i, DI_�������) = "" Then
                                    vsTmp.RemoveItem i
                                     j = vsTmp.Rows - 1
                                Else
                                    vsTmp.RowData(i) = 0
                                    i = i + 1
                                End If
                            Loop
                            '���ÿؼ���ʼ�Ŀ���״̬
                            Call SetDiagReletedInfo(vsTmp)
                            Call ChangeOutInfo
                        '���������й̶��еı�Ϊ�������Щ�����ֻ��������ݼ���
                        Else
                            vsTmp.Cell(flexcpData, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = Empty
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = ""
                            For i = vsTmp.FixedRows To vsTmp.Rows - 1
                                vsTmp.RowData(i) = 0
                            Next
                        End If
                    '�̶��е���0���̶��в�����0�ı���У����������ã����ƣ����ƣ�����ҩ����֢�໤������������Ŀ������ҩ
                    '��Щ���ֻ��Ҫɾ�����е��У�������һ�м��ɣ���֢�໤��������������֢�໤��Ҫ���������
                    ElseIf vsTmp.FixedRows <> 0 Then
                        If vsTmp.Name = "vsfMain" Then
                            For i = vsTmp.FixedRows To vsTmp.Rows - 1
                                For j = 0 To vsTmp.Cols - 1 Step 3
                                    If vsTmp.TextMatrix(i, j + 2) = "�Ƿ�" Then
                                        vsTmp.Cell(flexcpChecked, i, j + 1) = 2
                                    Else
                                        vsTmp.TextMatrix(i, j + 1) = ""
                                    End If
                                Next
                            Next
                        Else
                            '�����ĺ����ɾ�����е��У�������һ��
                            vsTmp.Rows = vsTmp.FixedRows
                            vsTmp.Rows = vsTmp.Rows + 1
                        End If
                    End If
                Case "PatiAddress" '��ַ�ؼ��Ƚ���ֻ��4�����ں���
                    Set paTmp = ctlTmp
                    '�ָ�Ĭ��ֵ
                    paTmp.Value = paTmp.Tag
                Case "ListBox" '��3��
                    Set lstTmp = ctlTmp
                    For i = 0 To lstTmp.ListCount - 1
                        lstTmp.Selected(i) = False
                    Next
                Case "ListView" 'lvwFees
                    Set lvwTmp = ctlTmp
                    For i = 0 To lvwTmp.ListItems.Count - 1
                        lvwTmp.ListItems(i).Checked = False
                    Next
                Case "MaskEdBox"
                    Set mskTmp = ctlTmp
                    If mskTmp.Index = DC_�ջ����� And gclsPros.OpenMode <> EM_���� Then
                    '������ҳ��Ŀ��������ʹ����ͬ���ջ�����
                    Else
                        mskTmp.Text = Replace(mskTmp.Tag, "#", "_")
                    End If
                Case "OptionButton"
                    Set optTmp = ctlTmp
                    optTmp.Value = (optTmp.Tag = "1")
                Case "VScrollBar"
                    Set vsbTmp = ctlTmp
                    vsbTmp.Value = 0
                Case "HScrollBar"
                    Set hsbTmp = ctlTmp
                    hsbTmp.Value = 0
            End Select
        Next
        Call SetFaceInit(True, True)
    End If
    gclsPros.IsOK = False
    If Not gclsPros.PatiInfo Is Nothing Then Set gclsPros.PatiInfo = zlDatabase.CopyNewRec(gclsPros.PatiInfo, True)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetFaceInit(Optional ByVal blnUnlock As Boolean, Optional ByVal blnReSetDefault As Boolean)
'���ܣ�������ظ�����ʼ״̬
' ���ܣ�blnUnlock=�Ƿ���ǩ������
'          blnReSetDefault=�Ƿ���������Ĭ��ֵ
    Dim objControl As Object
    Dim i As Long
    Dim LngRow As Long
    Dim strTmp As String, blnTmp As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim datCur  As Date
'    If blnReSetDefault Then Stop
    On Error GoTo errH
    With gclsPros.CurrentForm
        '�������пؼ�
        If blnUnlock Then
            For Each objControl In .Controls
                If InStr(",Timer,CommonDialog,Menu,Label,Subclass,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" Or TypeName(objControl.Container) = "Frame" Then
                            If Not (objControl.Name = "cmdSign") Then
                                Call SetCtrlLocked(objControl, False)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If gclsPros.FuncType <> f������ҳ Then
            '�����ض��ؼ���״̬
            Call SetCtrlLocked(.txtInfo(GC_����), True)
            Call SetCtrlLocked(.cboBaseInfo(BCC_�Ա�), True)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_����), True)
            Call SetCtrlLocked(.mskDateInfo(DC_��������), True)
        ElseIf gclsPros.FuncType = f������ҳ Then
            Call SetCtrlLocked(.txtInfo(GC_������), True)
        End If
        If gclsPros.PatiType = PF_סԺ Then
            strSql = "Select * From ���˱䶯��¼ Where ����ID=[1] And ��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", gclsPros.����ID, gclsPros.��ҳID)
            Call SetCtrlLocked(.txtInfo(BCC_��������), True)
            
            Call SetCtrlLocked(.mskDateInfo(DC_����ʱ��), Not IsDate(.mskDateInfo(DC_��������).Text))
            If gclsPros.FuncType <> f������ҳ Then
                Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), True)
                Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), True)
                Call SetCtrlLocked(.txtInfo(GC_��Ժ����), True)
                Call SetCtrlLocked(.txtInfo(GC_��Ժ����), True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_סԺ��), True)
            Else
                '��ZLHISϵͳ����������ѡ,���������޸ĵ���Ŀ������Ϊ��ɫ
                Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.txtInfo(GC_��Ժ����), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.txtInfo(GC_��Ժ����), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdDateInfo(DC_��Ժʱ��), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdDateInfo(DC_��Ժʱ��), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdInfo(GC_��Ժ����), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdInfo(GC_��Ժ����), IIf(rsTmp.RecordCount > 0, True, False), , True)
            End If
            Call SetCtrlLocked(.txtSpecificInfo(SLC_סԺ����), True)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_��Ժ����), True)
            Call SetDiagMatchInfo(BCC_�����벡��, True)
            Call SetDiagMatchInfo(BCC_�ٴ��벡��, True)
            Call SetDiagMatchInfo(BCC_�ٴ���ʬ��, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_�����), False)
            Call SetCtrlLocked(.chkInfo(CHK_��ԭѧ���), .vsDiagXY.TextMatrix(FindDiagRow(DT_Ժ�ڸ�Ⱦ), DI_�������) = "")
            LngRow = FindDiagRow(DT_��Ժ���XY)
            strTmp = UCase(Trim(.vsDiagXY.TextMatrix(LngRow, DI_��ϱ���)))
            blnTmp = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
            Call SetCtrlLocked(.cboBaseInfo(BCC_�ֻ��̶�), Not blnTmp, Not blnTmp)
            Call SetCtrlLocked(.cboBaseInfo(BCC_����������), Not blnTmp, Not blnTmp)

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_��Ժ��ʽ).Text)
            Call SetCtrlLocked(.mskDateInfo(DC_����ʱ��), strTmp <> "����")
            Call SetCtrlLocked(.txtInfo(GC_����ԭ��), strTmp <> "����")
            Call SetCtrlLocked(.cmdInfo(GC_����ԭ��), strTmp <> "����")
            Call SetCtrlLocked(.cboBaseInfo(BCC_��������ʬ��), strTmp <> "����")
            If .cboBaseInfo(BCC_��������ʬ��).ListIndex = -1 Then .cboBaseInfo(BCC_��������ʬ��).ListIndex = 0
            Call SetCtrlLocked(.chkInfo(CHK_����), strTmp = "����")
            Call SetCtrlLocked(.cboBaseInfo(BCC_�����ڼ�), strTmp <> "����")
            If gclsPros.FuncType = f������ҳ Then
                .cmdDeliceryInfo.Visible = False
                .cmdDeliceryInfo.Enabled = False
                .cmdDeliceryInfo.Tag = ""
                For i = LngRow To .vsDiagXY.Rows - 1
                    If Val(.vsDiagXY.TextMatrix(i, DI_��Ϸ���)) = DT_��Ժ���XY Then
                        If .vsDiagXY.TextMatrix(i, DI_������Ϣ) = "1" Then
                            .cmdDeliceryInfo.Visible = True
                            .cmdDeliceryInfo.Enabled = True
                            .cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If

            Call chkInfoClick(CHK_��ԭѧ���)
            Call chkInfoClick(CHK_����)
            Call chkInfoClick(CHK_�Ƿ�ȷ��)

            Call SetCtrlLocked(.txtInfo(GC_31������סԺ), .optInput(OP_��סԺ��).Value)
            Call SetCtrlLocked(.txtInfo(GC_���Ȳ���), Val(.txtSpecificInfo(SLC_���ȴ���).Text) = 0)
            Call SetCtrlLocked(.cmdInfo(GC_���Ȳ���), Val(.txtSpecificInfo(SLC_���ȴ���).Text) = 0)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_�ɹ�����), Val(.txtSpecificInfo(SLC_���ȴ���).Text) = 0)
            For i = 0 To .lstAdvEvent.ListCount - 1
                If .lstAdvEvent.List(i) = "ѹ��" Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_ѹ�������ڼ�), Not .lstAdvEvent.Selected(i))
                    Call SetCtrlLocked(.cboBaseInfo(BCC_ѹ������), Not .lstAdvEvent.Selected(i))
                ElseIf .lstAdvEvent.List(i) = "ҽԺ�ڵ���/׹��" Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_������׹���˺�), Not .lstAdvEvent.Selected(i))
                    Call SetCtrlLocked(.cboBaseInfo(BCC_������׹��ԭ��), Not .lstAdvEvent.Selected(i))
                End If
            Next

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_��Ժ��ʽ).Text)
            blnTmp = Not (strTmp Like "*תԺ*" Or strTmp Like "*ת����*")
            Call SetCtrlLocked(.txtInfo(GC_ת��ҽ�ƻ���), blnTmp)
            Call SetCtrlLocked(.cmdInfo(GC_ת��ҽ�ƻ���), blnTmp)

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_��Ժ;��).Text)
            blnTmp = Not (strTmp Like "*ת��*" And Not strTmp Like "*��ת��*")
            Call SetCtrlLocked(.txtInfo(GC_��Ժת��), blnTmp)
            Call SetCtrlLocked(.cmdInfo(GC_��Ժת��), blnTmp)

            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                blnTmp = zlStr.NeedName(.cboBaseInfo(BCC_��Һ��Ӧ).Text) <> "��"
                Call SetCtrlLocked(.txtInfo(GC_����ҩ��), blnTmp)
                Call SetCtrlLocked(.txtInfo(GC_�ٴ�����), blnTmp)
                Call chkInfoClick(CHK_�������)
            ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                blnTmp = .txtInfo(GC_��֢�໤������).Text = ""
                Call SetCtrlLocked(.chkInfo(CHK_�˹������ѳ�), blnTmp)
                Call SetCtrlLocked(.chkInfo(CHK_�ط���֢ҽѧ��), blnTmp)
                Call SetCtrlLocked(.cboBaseInfo(BCC_�ط����ʱ��), blnTmp)
                Call chkInfoClick(CHK_סԺ����Լ��)
            ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��֢�໤��), .optInput(OP_ICU��).Value)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��֢�໤Сʱ), .optInput(OP_ICU��).Value)
            End If
            If gclsPros.MedPageSandard = ST_����ʡ��׼ Or gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Call chkInfoClick(CHK_����·��)
                Call chkInfoClick(CHK_����)
                Call chkInfoClick(CHK_���·��)
            End If
            If gclsPros.FuncType = f������ҳ Then
                '������ҳ���ٱ༭����
                .lblNote.Visible = Not gclsPros.EditUnrecive
                .txtInfo(GC_������).TabStop = gclsPros.EditPageNo And gclsPros.OpenMode = EM_��������
                .txtInfo(GC_������).TabStop = gclsPros.TabFileNo
                .cboBaseInfo(BCC_���ʽ).TabStop = gclsPros.TabPayType
                .cboSpecificInfo(SLC_����).TabStop = gclsPros.TabAgeUnit
                .cboBaseInfo(BCC_����).TabStop = gclsPros.TabNation
                .txtInfo(GC_������).Locked = Not gclsPros.EditPageNo
                .chkInfo(CHK_����Ժ).TabStop = gclsPros.TabReadm
                .txtInfo(GC_X�ߺ�).TabStop = gclsPros.TabXRaysNo

                If gclsPros.OpenMode = EM_���� Or gclsPros.OpenMode = EM_�༭ Then
                    frmMain.cbsMain.FindControl(, conMenu_Manage_Up, True).Enabled = Get��ҳIDByCur(gclsPros.��ҳID, False) <> 0
                    frmMain.cbsMain.FindControl(, conMenu_Manage_Down, True).Enabled = Get��ҳIDByCur(gclsPros.��ҳID, True) <> 0
                End If

                Call SetCtrlLocked(.txtSpecificInfo(SLC_סԺ��), gclsPros.OpenMode <> EM_�������� And gclsPros.OpenMode <> EM_������ҳ)
                Call SetCtrlLocked(.vsFees, True)
                Call SetCtrlLocked(.cboManInfo(MC_��ĿԱ), Not gclsPros.Change����Ա)
            Else
                Call SetCtrlLocked(.cboBaseInfo(BCC_���ʽ), InStr(gclsPros.Privs, "�޸�ҽ�Ƹ��ʽ") = 0)
            End If
            '�Ĵ����ȡ�ϴ���ϡ��������ϻ�ȡ�ϴ����
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ And gclsPros.FuncType = f������ҳ Then
                .cmdLastDiag.Visible = gclsPros.��ҳID > 1
                .lblDiagInfo.Caption = ""
                .lblDiagInfo.Visible = False
            End If
        Else
            '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
            blnTmp = IsDate(.vsDiagXY.TextMatrix(.vsDiagXY.FixedRows, DI_����ʱ��)) Or IsDate(.vsDiagZY.TextMatrix(.vsDiagZY.FixedRows, DI_����ʱ��))
            Call SetCtrlLocked(.mskDateInfo(DC_��������), blnTmp)
            Call SetCtrlLocked(.mskDateInfo(DC_����ʱ��), blnTmp)
            Call SetCtrlLocked(.cboBaseInfo(BCC_���ʽ), InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") = 0)
            '���õ�λ���ƵĿ�������
            blnTmp = InStr(gclsPros.Privs, "��Լ���˵Ǽ�") = 0 And Not IsNull(gclsPros.PatiInfo!��ͬ��λid)
            Call SetCtrlLocked(.txtAdressInfo(ADRC_��λ��ַ), blnTmp)
        End If
        If .cboBaseInfo(BCC_Ѫ��).ListIndex = -1 Then .cboBaseInfo(BCC_Ѫ��).ListIndex = 0
        If .cboBaseInfo(BCC_RH).ListIndex = -1 Then .cboBaseInfo(BCC_RH).ListIndex = 0
        '�ָ�Ĭ��ֵ
        If blnReSetDefault Then
            Call SetCboDefault(.cboBaseInfo(BCC_���֤), -1)
            Call SetCboDefault(.cboBaseInfo(BCC_����״��), 0)
            Call SetCboDefault(.cboSpecificInfo(SLC_����), 0)
            If gclsPros.PatiType <> PF_���� Then
                Call SetCboDefault(.cboSpecificInfo(SLC_��������), 0)
                Call SetCboDefault(.cboSpecificInfo(SLC_Ӥ�׶�����), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_����Ժ�ƻ�����), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_��Ⱦ��������ϵ), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_��������ʬ��), 0)
                If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    Call SetCboDefault(.cboBaseInfo(BCC_�ٴ�·������), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_ʵʩDGRS����), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_������Ⱦ��), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_��������), 0)
                End If
            End If
            '����һЩ�ֵ���������������
            Call SetCboDefaultByRec(Array(BCC_���ʽ, BCC_�Ա�, BCC_����, BCC_ְҵ, BCC_����, BCC_����, BCC_Ѫ��))
            If gclsPros.PatiType <> PF_���� Then
                Call SetCboDefaultByRec(Array(BCC_��������, BCC_��ϵ, BCC_��Ժ���, BCC_��Ժ;��, BCC_�ֻ��̶�, BCC_����������, BCC_��Ժ��ʽ))
            Else
                Call SetCboDefaultByRec(Array(BCC_ȥ��, BCC_�Ļ��̶�))
            End If
            Call SetCboDefaultByRec(Array(BCC_�����ڼ�))
            If gclsPros.FuncType = f������ҳ Then
                '�õ�Ĭ�ϳ�����
                strSql = "select A.����,A.���� from ���� a where a.ȱʡ��־=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_�����ص�, "�����ص�", rsTmp!����, True)
                    If gclsPros.DefautADD Then
                        Call SetPatiAddress(ADRC_��ϵ�˵�ַ, "��ϵ�˵�ַ", rsTmp!����, True)
                        Call SetPatiAddress(ADRC_��סַ, "��ͥ��ַ", rsTmp!����, True)
                        .txtSpecificInfo(SLC_��ͥ�ʱ�).Text = rsTmp!���� & ""
                    End If
                End If
                '����:13557
                strSql = "select A.����,A.���� from ���� a where a.ȱʡ��־=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_��������, "����", rsTmp!����, True)
                End If
                datCur = zlDatabase.Currentdate
                '����Ĭ��ֵ
                .mskDateInfo(DC_��Ŀ����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ŀ����).Tag))
                If Not IsDate(.mskDateInfo(DC_�ջ�����).Text) Then
                    .mskDateInfo(DC_�ջ�����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_�ջ�����).Tag))
                End If
                .mskDateInfo(DC_��������).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��������).Tag))
                .mskDateInfo(DC_��Ժʱ��).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ժʱ��).Tag))
                .mskDateInfo(DC_��Ժʱ��).Text = Format(datCur, GetFormat(.mskDateInfo(DC_��Ժʱ��).Tag))
                .mskDateInfo(DC_�ʿ�����).Text = Format(datCur, GetFormat(.mskDateInfo(DC_�ʿ�����).Tag))

                .txtDateInfo(DC_��Ŀ����).Text = .mskDateInfo(DC_��Ŀ����).Text
                .txtDateInfo(DC_�ջ�����).Text = .mskDateInfo(DC_�ջ�����).Text
                .txtDateInfo(DC_��������).Text = .mskDateInfo(DC_��������).Text
                .txtDateInfo(DC_��Ժʱ��).Text = .mskDateInfo(DC_��Ժʱ��).Text
                .txtDateInfo(DC_��Ժʱ��).Text = .mskDateInfo(DC_��Ժʱ��).Text
                .txtDateInfo(DC_�ʿ�����).Text = .mskDateInfo(DC_�ʿ�����).Text

                .cboManInfo(MC_��ĿԱ).Text = UserInfo.����
                gclsPros.InTime = .mskDateInfo(DC_��Ժʱ��).Text
                gclsPros.OutTime = .mskDateInfo(DC_��Ժʱ��).Text
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFaceEditable(ByVal blnSign As Boolean)
'���ܣ����ý���ؼ�������
'������blnSign=�Ƿ���ǩ�����ý��棬True-����ǩ��״̬���ã�False-������ǩ��״̬
    Dim objControl As Object
    Dim blnEnable As Boolean
    Dim bln������Ϣ As Boolean
    Dim lngTabIndex��Ժ As Long
    Dim strTypeName As String
    Dim lngMainTab As Long
    Dim blnSet As Boolean

    On Error GoTo errH
    With gclsPros.CurrentForm
        '����״̬������ؼ���������
        If gclsPros.OpenMode = EM_���� Then
            For Each objControl In .Controls
                If InStr(",Timer,CommonDialog,Menu,Label,Subclass,VScrollBar,HScrollBar,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" Or TypeName(objControl.Container) = "Frame" Then
                            If objControl.Name = "cmdDeliceryInfo" Then '����鿴��ť
                                Call SetCtrlLocked(objControl, objControl.Tag = "")
                            Else
                                Call SetCtrlLocked(objControl, True)
                            End If
                        End If
                    End If
                End If
            Next
        '�༭״̬���ݾ�������������
        Else
            '���л�����Ϣ��Ȩ�޵ģ���Ժʱ��֮ǰ�Ŀؼ���û����д���������д�����򲻿ɱ༭�����˺��������Ա����䡢�������ڲ��ɱ༭��ֻ��������ȼ��Լ����˻�����Ϣ�޸����޸ģ�
            If gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiType = PF_סԺ And Not gclsPros.Is��ʿվ Then
                lngTabIndex��Ժ = .lblBaseInfo(BCC_��Ժ;��).TabIndex
                bln������Ϣ = InStr(";" & gclsPros.Privs & ";", ";��ҳ������Ϣ;") > 0
            Else
                lngTabIndex��Ժ = 0
                bln������Ϣ = True
            End If
            For Each objControl In .Controls
                strTypeName = TypeName(objControl): blnSet = True: blnEnable = Not blnSign
                '��Ҫ�жϵ����
                If InStr(",Timer,CommonDialog,Menu,Frame,Label,Line,Subclass,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" And InStr(",PicPage,PicMain,", "," & objControl.Name & ",") = 0 Or TypeName(objControl.Container) = "Frame" Then
                            If blnEnable Then
                                'ҽ���뻤ʿ������ҳ�Ŀ���
                                If objControl.Container.Name = "PicAdvEvent" Or objControl.Container.Name = "PicRestrain" Or objControl.Container.Name = "PicCareInfo" Then
                                    blnEnable = Not gclsPros.SeparateEdit Or gclsPros.Is��ʿվ And gclsPros.SeparateEdit
                                Else
                                    blnEnable = Not gclsPros.SeparateEdit Or Not gclsPros.Is��ʿվ And gclsPros.SeparateEdit
                                End If
                                '��ҳ������Ϣ����
                                If blnEnable Then
                                    If gclsPros.FuncType = fҽ����ҳ And objControl.TabIndex < lngTabIndex��Ժ Then
                                        blnSet = Not ControlIsLocked(objControl)
                                        blnEnable = bln������Ϣ Or Not ControlHaveValue(objControl)
                                    ElseIf gclsPros.PatiType = PF_סԺ And blnEnable Then
                                        '���в�����Ժ���Ժ���Ҿ�����ҽ�����ʣ����Ҳ���"��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ"=True��
                                        '�������͡���Һ��Ӧ����Ѫ��Ӧ�����ϸ������ѪС�塢��Ѫ������ȫѪ��������ա�����������Ѫǰ��9���顢
                                        'HBsAg��HCV-Ab��HIV-Ab��ʾ�̲��������в���������������ޡ�������ʹ�á��о���ҽʦ.���ò�����
                                        If gclsPros.Have��ҽ And gclsPros.NotUseXYItems Then
                                            Select Case objControl.Name
                                                Case "cboBaseInfo"
                                                    blnEnable = Not (objControl.Index = BCC_�������� Or objControl.Index = BCC_��Һ��Ӧ Or objControl.Index = BCC_��Ѫ��Ӧ Or _
                                                                objControl.Index = BCC_��Ѫǰ9���� Or objControl.Index = BCC_HBsAg Or objControl.Index = BCC_HCVAb Or _
                                                                objControl.Index = BCC_HIVAb)
                                                Case "txtSpecificInfo"
                                                    blnEnable = Not (objControl.Index = SLC_���ϸ�� Or objControl.Index = SLC_��ȫѪ Or objControl.Index = SLC_��Ѫ�� Or _
                                                                objControl.Index = SLC_������� Or objControl.Index = SLC_��ѪС�� Or objControl.Index = SLC_������ʹ�� Or _
                                                                objControl.Index = SLC_�������� Or objControl.Index = SLC_��׵���)
                                                Case "cboManInfo"
                                                    blnEnable = objControl.Index <> MC_�о���ҽʦ
                                                Case "chkInfo"
                                                    blnEnable = Not (objControl.Index = CHK_���в��� Or objControl.Index = CHK_ʾ�̲��� Or objControl.Index = CHK_����)
                                                Case "txtInfo"
                                                    blnEnable = objControl.Index <> GC_������
                                                Case "cboSpecificInfo"
                                                    blnEnable = objControl.Index <> SLC_��������
                                            End Select
                                        End If
                                        blnSet = Not ControlIsLocked(objControl)
                                    Else
                                        blnSet = Not ControlIsLocked(objControl)
                                    End If
                                End If
                            End If
                            If objControl.Name = "cmdSign" Or objControl.Name = "cmdUnSign" Then
                                blnSet = gclsPros.Is��ʿվ
                            End If
                            If blnSet Then
                                Call SetCtrlLocked(objControl, Not blnEnable)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetMedSt(ByVal strCurFormType As String) As MedPage_Standard
'���ܣ���ȡ��ҳ��׼
    Dim strFixed As String, strOther As String
    If strCurFormType = "" Then Exit Function
    strFixed = decode(gclsPros.FuncType, fҽ����ҳ, "frmInMedRecEdit", f������ҳ, "frmPageMedRecEdit", f���Ӳ���, "frmArchiveInMedRec", "")
    strOther = Replace(strCurFormType, strFixed, "")
    Select Case strOther
        Case ""
            GetMedSt = ST_��������׼
        Case "_SC"
            GetMedSt = ST_�Ĵ�ʡ��׼
        Case "_YN"
            GetMedSt = ST_����ʡ��׼
        Case "_HN"
            GetMedSt = ST_����ʡ��׼
        Case "frmOutMedRecEdit", "frmArchiveOutMedRec"
            GetMedSt = ST_������ҳ
    End Select
End Function

Public Sub SavePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSql As String

    On Error GoTo Errhand
    'û�иı�ͼƬ������
    If gclsPros.CurrentForm.picPatient.Tag = gclsPros.PictureFile Then gclsPros.PictureFile = "0": Exit Sub
    gclsPros.PictureFile = ""
    'ͼƬû�б�����������²���ͼƬ
    If gclsPros.CurrentForm.picPatient.Tag <> "" Then
        strFile = gclsPros.CurrentForm.picPatient.Tag
        If sys.SaveLob(gclsPros.SysNo, 27, lng����ID, strFile) = False Then
            MsgBox "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    gclsPros.PictureFile = gclsPros.CurrentForm.picPatient.Tag
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub VsGriedFocuesMove(ByRef vsBill As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal KeyCode As Integer, _
        Optional lngFiexCol As Long = 0, Optional lngFiexCol1 As Long = -1)
    '------------------------------------------------------------------------------------------------------------
    '����:��һ�������ƶ���Ԫ��
    '����:vsBill-���ؼ�
    '       lngRow-��ǰ��
    '       lngCol-��ǰ��
    '       KeyCode-����
    '       lngFiexCol-�ж��Ƿ��Ƶ�������еĹ̶���
    '       lngFiexCol1-�ж��Ƿ��Ƶ�������еĹ̶���(��ͬʱҪ����lngFiexCol��)
    '����:���˺�
    '����:2007/05/18
    '------------------------------------------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim strCurrValue As String, strTmp As String, LngCols As Long, i As Long
    If LngCol = lngFiexCol Then
        strCurrValue = vsBill.EditText
    Else
        strCurrValue = ""
    End If

    With vsBill
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then LngCols = LngCols + 1
        Next
        Select Case LngCol
        Case 0
            If Trim(.TextMatrix(LngRow, lngFiexCol)) = "" And strCurrValue = "" And vsBill.Name <> "vsFlxAddICU" Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            .Col = LngCol + 1
            GoTo ShowCell:
        Case Else
            If LngCol >= LngCols - 1 Then
                If LngRow < .Rows - 1 Then
                    .Row = LngRow + 1
                    .Col = .FixedCols
                    GoTo ShowCell:
                    Exit Sub
                End If
                If vsBill.Name = "vsFlxAddICU" Then
                    strTmp = Trim(.TextMatrix(LngRow, lngFiexCol))
                    strTmp = Replace(strTmp, ":", "")
                    strTmp = Replace(strTmp, "-", "")
                    strTmp = Replace(strTmp, "_", "")
                    strTmp = Replace(strTmp, " ", "")
                Else
                    strTmp = Trim(.TextMatrix(LngRow, lngFiexCol))
                End If
                If strTmp <> "" Then
                    If lngFiexCol1 > 0 Then
                        If Trim(.TextMatrix(LngRow, lngFiexCol1)) <> "" Then
                            .Rows = .Rows + 1
                            Call ChangeVSFHeight(vsBill, True)
                            .Row = .Rows - 1
                            .Col = .FixedCols
                        End If
                    Else
                        .Rows = .Rows + 1
                        Call ChangeVSFHeight(vsBill, True)
                        .Row = .Rows - 1
                        .Col = .FixedCols
                    End If
                Else
                    zlCommFun.PressKey vbKeyTab
                    Exit Sub
                End If
                GoTo ShowCell:
                Exit Sub
            End If
            .Col = LngCol + 1
         End Select
ShowCell:
        .ShowCell .Row, .Col
    End With
End Sub

Public Function LoadPatiByInNo(ByVal strסԺ�� As String, Optional ByVal lng��ҳID As Long, Optional ByVal str������ As String) As Boolean
'���ݵ�ǰסԺ�ż��ز�����Ϣ
    Dim blnOut As Boolean '�Ƿ��ⲿ�ļ���ȡ
    Dim lng���� As Long, lng�ϴδ��� As Long, blnNoCheck As Boolean
    Dim blnOrderAdd As Boolean 'Ԥ������:�����Ƿ�ֻ�ܰ�˳�����(ĿǰȱʡFalse,�Ա��Ժ���չʹ��)
    Dim lngTemp As Long
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim rs������ҳ As ADODB.Recordset
    Dim intSel��ҳid As Integer
    Dim blnNO��ҳID As Boolean
    If strסԺ�� <> gclsPros.InNo Then
        Call ClearPageContent
    End If
    gclsPros.IsExistPati = False
    gclsPros.Is��Ŀ = False
    If strסԺ�� = "" Then
        '63725:������,2013-08-06
        If Not gclsPros.EditUnrecive Then Exit Function
        If gclsPros.OnLine Then
            '��������Ҳ�����������ֻ�ܴ����еĳ�Ժ�����еõ�
            If Not gclsPros.OnLineNew Then Exit Function
            MsgBox "��ǰ��������һλ���շ�ϵͳ�����ڵĲ��ˡ�", vbInformation, gstrSysName
            '��ZLHISϵͳ����������ѡ,���������޸ĵ���Ŀ������Ϊ��ɫ
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
        End If
        gclsPros.OpenMode = EM_��������
        '��ȡ�����Ĳ��˵ĸ��ֺ���
        Call ValidatePageNos
        If Not ExistInList(gclsPros.InNo, True) Then Exit Function
    Else
        gclsPros.InNo = strסԺ��
        If Not ExistInList(gclsPros.InNo, True) Then Exit Function
        If Not IsHavePageNos(CT_סԺ��, False, gclsPros.InNo) Then
            If Not gclsPros.EditUnrecive Then
                MsgBox "��סԺ�����շ�ϵͳ�в�����,���ܼ�����", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                Exit Function
            End If
            If gclsPros.OnLine Then
                '��������Ҳ�����������ֻ�ܴ����еĳ�Ժ�����еõ�
                If Not gclsPros.OnLineNew Then
                    MsgBox "��סԺ�����շ�ϵͳ�в���,���ܼ�����", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                    Exit Function
                Else
                    If Not gclsPros.NewInNo Then
                        '������Ϣ���Ƿ���ڸ�סԺ��
                        If IsHavePageNos(CT_סԺ��ex, False, gclsPros.InNo) Then
                            If MsgBox("סԺ��Ϊ" & gclsPros.InNo & "�Ĳ����ڲ�����ҳ��������Ϣ���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                                Exit Function
                            End If
                        Else
                            If MsgBox("סԺ��Ϊ" & strסԺ�� & "�Ĳ�����ϵͳ�в������κ���Ϣ���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                                Exit Function
                            End If
                        End If
                        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                    End If
                End If
            End If
            'ͨ���ļ����ⲿ���ݿ�õ�������Ϣ
            If gclsPros.OutFile <> "" Then
                gclsPros.PatiOut.Filter = "סԺ��= " & IIf(strסԺ�� = "", 0, strסԺ��) & IIf(lng��ҳID = 0, "", " and סԺ����=" & lng��ҳID)
                If gclsPros.PatiOut.EOF Then
                    If MsgBox("סԺ��Ϊ" & strסԺ�� & "��סԺ����Ϊ1�Ĳ������ⲿ�ļ���û�ҵ����Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                        zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                        gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                        Exit Function
                    End If
                    blnOut = False
                Else
                    blnOut = True
                End If
            End If
            If blnOut Then
                gclsPros.��ҳID = lng��ҳID
                If gclsPros.IsSelPati Then Call LoadDataFromOutFile(strסԺ��)
            End If
            gclsPros.NoType = IT_New
            If gclsPros.NewInNo Then
                If gclsPros.OpenMode <> EM_������ҳ Then
                    strTemp = zlCommFun.ShowMsgbox(gclsPros.CurrentForm.Caption, "סԺ��Ϊ��" & strסԺ�� & "���Ĳ���δ�ҵ�����ȷ��������ʽ��", "!��������(&A),������ҳ(&N)", gclsPros.CurrentForm, vbQuestion)
                Else
                    strTemp = "��������"
                    gclsPros.OpenMode = EM_��������
                End If
                If strTemp = "��������" Then
                    gclsPros.OpenMode = EM_��������
                    gclsPros.CurrentForm.txtInfo(GC_������).Text = strסԺ��
                Else
                    gclsPros.OpenMode = EM_������ҳ
                    gclsPros.NoType = IT_NewMed
                    If gclsPros.EditPageNo Then
                        gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text
                    End If
                End If
                If gclsPros.OnLineNew Then
                    '��ZLHISϵͳ����������ѡ,���������޸ĵ���Ŀ������Ϊ��ɫ
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                End If
            End If
            gclsPros.OnlyPatiInfo = False
            If Not blnOut Then
                strSql = "select ����id from ������Ϣ where סԺ��=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strסԺ��)
                If rsTmp.RecordCount > 0 Then
                    gclsPros.����ID = Val(rsTmp!����ID & "") '���õ��򣬵�һ����¼�������һ��
                    gclsPros.��ҳID = 1
                    gclsPros.OnlyPatiInfo = True
                    lng��ҳID = IIf(gclsPros.SinPageNo, lng��ҳID, 0)
                    Call LoadMedPageData(gclsPros.����ID, IIf(gclsPros.SinPageNo, gclsPros.��ҳID, 0), , , gclsPros.Is��Ŀ)
                End If
            End If
            If gclsPros.NoType = IT_New And Not gclsPros.OnlyPatiInfo Then
                '��������,������ȷ������ID
                gclsPros.����ID = NVL(GetNextNo(1))    'סԺʹ���û��Լ�����ģ�������ID��Ҫ�Զ�����
                If Not blnOut Then gclsPros.��ҳID = 1
            End If
            If blnOut Then
                If gclsPros.OutFile <> "" And Not gclsPros.IsSelPati Then
                    gclsPros.��ҳID = Select�ⲿ��ҳid(strסԺ��, Val(lng��ҳID))
                End If
            End If
            gclsPros.CurrentForm.txtSpecificInfo(SLC_��Ժ����) = gclsPros.��ҳID
            If Not gclsPros.EditPageNo Or Trim(gclsPros.CurrentForm.txtInfo(GC_������).Text) = "" Then
                '���ܱ༭ ���� ������Ϊ�գ���ʱ�Զ��滻
                gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.InNo
            End If
            If gclsPros.EditPageNo Then
                '����༭�����ţ�����ͣ
                gclsPros.CurrentForm.txtInfo(GC_������).TabStop = True
            End If
            '53638:������,2013-05-10,���������ű�Ź���
            If gclsPros.UseFileRules Then
                gclsPros.CurrentForm.txtInfo(GC_������).Text = NVL(GetNextNo(CT_������, , GetDeptCode(gclsPros.��Ժ����ID)))
            End If

            If gclsPros.NoType = IT_New Then
                gclsPros.OpenMode = EM_��������
            Else
                gclsPros.OpenMode = EM_������ҳ
            End If
            gclsPros.Is��Ŀ = False
            LoadPatiByInNo = True
            Exit Function
        Else
            gclsPros.IsExistPati = True
        End If
        strSql = "select ����id from ������ҳ where סԺ��=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strסԺ��)
        gclsPros.����ID = Val(rsTmp!����ID & "")
        '�ó������ڲ���������ҳ��Ϣ
        strSql = "select ��ҳID from ������ҳ where ����ID=[1] and nvl(��������,0)=0 and ��Ŀ���� is not null  order by ��ҳID Desc"
        Set rs������ҳ = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, Val(gclsPros.����ID & ""))
        If rs������ҳ.RecordCount > 0 Then
            lng���� = rs������ҳ("��ҳID") '���õ��򣬵�һ����¼�������һ��
            lng�ϴδ��� = lng����
        End If

        If Not gclsPros.IsSelPati Then
            '78747:ȡ����ҳֻ������������ӵ�����
            If gclsPros.OutFile = "" Then
                intSel��ҳid = Select��ҳID(gclsPros.����ID, IIf(blnOrderAdd = True, lng����, 0))
            Else
                intSel��ҳid = Select�ⲿ��ҳid(strסԺ��, IIf(blnOrderAdd = True, lng����, 0))
            End If
            If lng���� > intSel��ҳid And intSel��ҳid <> 0 And blnOrderAdd = True Then
                ' ���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
                ' 2007/05/10
                Call GetסԺ����Or��ҳid(gclsPros.����ID, lng����, lngTemp, False)
                lngTemp = IIf(lngTemp = 0, lng����, lngTemp)
               '��סԺ����������ѡ���סԺ����ʱ�˳� ����2005-8-22
                MsgBox "��ѡ��ò����ڵ�" & lngTemp & "����Ժ�Ժ����Ϣ��", vbInformation, gstrSysName
                LoadPatiByInNo = False
                Exit Function
            ElseIf intSel��ҳid <> 0 Then
                '��ѡ���סԺ����>����������סԺ����ʱ������ѡ�����Ϣ��������
                lng���� = intSel��ҳid
                blnNoCheck = lng���� > intSel��ҳid
            ElseIf intSel��ҳid = 0 Then
                '63725:������,2013-08-06
                '���˺�:��Ϊ��������Ҫ���ƵĲ���,�����˳���.
                If (Not gclsPros.OnLineNew And gclsPros.OnLine And gclsPros.OutFile = "") Or Not gclsPros.EditUnrecive Then
                    Call GetסԺ����Or��ҳid(gclsPros.����ID, lng����, lngTemp, False)
                    lngTemp = IIf(lngTemp = 0, lng����, lngTemp)
                    MsgBox "�ò����ܹ�" & lngTemp & "��סԺ,�����Ѿ������˲���,���ܼ�����", vbInformation, gstrSysName
                    LoadPatiByInNo = False
                    Exit Function
                End If
            End If
        Else
           If gclsPros.OutFile <> "" Then
                If lng���� > lng��ҳID Then
                    ' ���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
                    ' 2007/05/10
                    Call GetסԺ����Or��ҳid(gclsPros.����ID, lng����, lngTemp, False)
                    lngTemp = IIf(lngTemp = 0, lng����, lngTemp)
                    MsgBox "��ѡ��ò����ڵ�" & lngTemp & "����Ժ�Ժ����Ϣ��", vbInformation, gstrSysName
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = "": gclsPros.InNo = ""
                    LoadPatiByInNo = False
                    Exit Function
                End If
           End If
           lng���� = lng��ҳID
        End If
        '�ж��ϴεĳ�Ժ���
        If lng�ϴδ��� > 0 And blnNoCheck = False Then
            strSql = "" & _
                "   Select B.��Ժ��� " & _
                "   From ������ҳ A,������ϼ�¼ B " & _
                "   Where A.����ID=[1] and A.��ҳID=[2] " & _
                "           and A.����ID = B.����ID And A.��ҳID = B.��ҳID And B.������� = 3 And B.��ϴ��� = 1 And B.������� = 1 " & _
                "           and a.��Ŀ���� is not null "
            Set rs������ҳ = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng�ϴδ���)

            If rs������ҳ.EOF = False Then
                If rs������ҳ("��Ժ���") = "����" Then
                    MsgBox "�ò��˵�" & lng�ϴδ��� & "�γ�Ժ����Ѿ����������������������ҳ�ˡ�", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                    Exit Function
                End If
            End If
        End If
        If gclsPros.OutFile <> "" Then
             If intSel��ҳid <> 0 Or gclsPros.IsSelPati Then
                gclsPros.��ҳID = lng����
             Else
                lng���� = lng���� + 1
             End If
        Else
            '������Ѿ���5����ҳ���룬���⽫�����ĵ�6��סԺ
            strSql = "" & _
                "   SELECT MIN(��ҳID) as סԺ���� " & _
                "   FROM ������ҳ " & _
                "   WHERE ����ID=[1] AND ��ҳID> =[2]" & _
                "  AND ��Ŀ���� IS NULL AND nvl(��������,0)=0" 'δ��Ŀ����Ϊ����סԺ

            Set rs������ҳ = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng����)
            If IsNull(rs������ҳ("סԺ����")) Then

                If gclsPros.OnLine Then
                    '����Ƿ�������۲���,�������,ȡ�������۲��˵���ҳID+1
                    strSql = "" & _
                        "   SELECT max(��ҳID) as סԺ���� " & _
                        "   FROM ������ҳ " & _
                        "   WHERE ����ID=[1] AND ��ҳID> =[2]" & _
                        "           AND nvl(��������,0)<>0" '����Ŀ���ڵ����۲��˵���ҳID
                    Set rs������ҳ = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng����)

                    If rs������ҳ.EOF Then
                        '֤�������ڴ�סԺ����,��ȡ����סԺ����
                        lng���� = lng���� + 1
                    Else
                        If Val(NVL(rs������ҳ("סԺ����"))) = 0 Then
                            lng���� = lng���� + 1
                        Else
                            '֤���������۲���,�����Ҫ���������۲��˵���ҳid+1
                            lng���� = Val(NVL(rs������ҳ("סԺ����"))) + 1
                        End If
                    End If
                Else
                    lng���� = lng���� + 1
                End If
            Else
                lng���� = rs������ҳ("סԺ����") '�����м��������۲��ˣ����Բ���ֱ��ȡ���ֵ+1
            End If
        End If

        '�ó�������Ժ�շѴ�����ҳ��Ϣ
        strSql = "Select ��Ժ����, ��Ŀ���� from ������ҳ where ����ID=[1] and ��ҳID= [2] And nvl(��������,0)=0"
        Set rs������ҳ = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng����)
        If rs������ҳ.RecordCount <> 0 Then
            If IsNull(rs������ҳ("��Ժ����")) Then
                MsgBox "�ò�����Ȼ��Ժ��������д��ҳ��", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                Exit Function
            End If
        End If
        '63725:������,2013-08-06
        If Not gclsPros.EditUnrecive Then
            strSql = "Select ID from �������ռ�¼ Where ����ID=[1] and ��ҳID= [2] And ����ʱ�� IS NOT NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng����)
            If rsTmp.RecordCount = 0 Then
                MsgBox "��ǰҪ��Ŀ���ǲ��˵�" & lng���� & "��סԺ���������һ�û�н��գ����ܽ��б�Ŀ����!", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                Exit Function
            End If
        End If

        If lng���� > 1 Then
            ' ���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
            ' 2007/05/10
            lngTemp = 0
            If GetסԺ����Or��ҳid(gclsPros.����ID, lng����, lngTemp, False) = False Then
                MsgBox "��ȡָ����ҳ�Ĵ���ʧ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            lngTemp = IIf(lngTemp = 0, lng����, lngTemp)
            If MsgBox("���������벡�˵ĵ�" & lngTemp & "����ҳ���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                Exit Function
            End If
        End If
        blnNO��ҳID = False
        If rs������ҳ.RecordCount = 0 Then
            If gclsPros.OnLine Then
                 ' ���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
                ' 2007/05/10
                Call GetסԺ����Or��ҳid(gclsPros.����ID, lng����, lngTemp, False)
                lngTemp = IIf(lngTemp = 0, lng����, lngTemp)
                '����32713 by lesfeng 2010-09-13
                If Not gclsPros.OnLineNew Then
                     MsgBox "�ò��˵ĵ�" & lngTemp & "��סԺ��Ϣ���շ�ϵͳ��û�ҵ������ܼ�����", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                    Exit Function
                Else
                    MsgBox "�ò��˵ĵ�" & lngTemp & "��סԺ��Ϣ���շ�ϵͳ��û�ҵ���Ŀǰ�������շ�ϵͳ�в����ڵ���ҳ��", vbInformation, gstrSysName
                    '��ZLHISϵͳ����������ѡ,���������޸ĵ���Ŀ������Ϊ��ɫ
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_��Ժʱ��), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_��Ժ����), False)
                End If
            Else
                If gclsPros.OutFile <> "" Then
                    'ͨ���ļ����ⲿ���ݿ�õ�������Ϣ
                    gclsPros.PatiOut.Filter = "סԺ��= " & IIf(strסԺ�� = "", 0, strסԺ��) & IIf(lng���� = 0, "", " and סԺ����=" & lng����)
                    If gclsPros.PatiOut.EOF Then
                        If MsgBox("סԺ��Ϊ" & strסԺ�� & "�Ĳ������ⲿ�ļ���û�ҵ����Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                            zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                            Exit Function
                        End If
                    End If
                    blnOut = True
                End If
                blnNO��ҳID = True
            End If
            '��Ϊ������ҳģʽ
            gclsPros.OpenMode = EM_������ҳ
            If gclsPros.NewInNo Then
                '���˺�:��Ҫ�²���һ��סԺ��
                gclsPros.InNo = NVL(GetNextNo(2))
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = gclsPros.InNo
            End If
            Call LoadDataFromPaitiInfo(gclsPros.����ID, lng����)
            If blnOut Then
                '���˺�
                gclsPros.��ҳID = lng����
                Call LoadDataFromOutFile(strסԺ��)
            End If
            gclsPros.Is��Ŀ = False
        Else
            If IsNull(rs������ҳ("��Ժ����")) Then
                MsgBox "�ò�����Ȼ��Ժ��������д��ҳ��", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).SetFocus
                Exit Function
            End If
            Call LoadDataFromPaitiInfo(gclsPros.����ID, lng����)
            strSql = "Select 1 From ������ϼ�¼ Where ����ID=[1] And ��ҳID=[2] And ��¼��Դ=4 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, lng����)
            '��Ϊ������ҳģʽ����ֻ�Ƕ�������ҳ���б�Ŀ
            gclsPros.OpenMode = EM_������ҳ
            gclsPros.Is��Ŀ = True
            Call LoadMedPageData(gclsPros.����ID, lng����, , , rsTmp.RecordCount > 0)
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Or gclsPros.MedPageSandard = ST_����ʡ��׼ And gclsPros.FuncType = f������ҳ Then
                gclsPros.CurrentForm.cmdLastDiag.Visible = lng���� > 1
                gclsPros.CurrentForm.lblDiagInfo.Caption = ""
                gclsPros.CurrentForm.lblDiagInfo.Visible = False
            End If
            Call SetPageVisible
            Call SetPicPosition(True)
        End If
        '0-��ǰ���˵�סԺ����û������֤�ģ�1-סԺ�����µģ�2-סԺ������ǰ��:3-סԺ����������ҳ��
        gclsPros.NoType = IT_Old
        gclsPros.��ҳID = lng����
        Call ValidatePageNos
        lngTemp = 0
        If blnNO��ҳID Then
        Else
            ' ���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
            ' 2007/05/10
            Call GetסԺ����Or��ҳid(gclsPros.����ID, gclsPros.��ҳID, lngTemp, False)
            If intSel��ҳid = 0 And gclsPros.��ҳID <> 1 And Not gclsPros.IsSelPati Then
                lngTemp = lngTemp + 1
            End If
        End If
        gclsPros.CurrentForm.txtSpecificInfo(SLC_��Ժ����).Text = IIf(lngTemp = 0, lng����, lngTemp)
    End If
    LoadPatiByInNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub AfterLoadPatiByNo()
    Dim blnEditInfo As Boolean

    Call SetFaceEditable(gclsPros.IsSigned)
    If gclsPros.OpenMode <> EM_�༭ And Not gclsPros.Is��Ŀ Then
        blnEditInfo = True '���Ӳ�����ҳ
    End If
    If gclsPros.NoType = IT_New Then
        blnEditInfo = True '���Ӳ�����ҳ
    End If
    With gclsPros.CurrentForm
        '��ZLHISϵͳ����������ѡ,���������޸ĵ���Ŀ������Ϊ��ɫ
        Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.mskDateInfo(DC_��Ժʱ��), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.txtInfo(GC_��Ժ����), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.txtInfo(GC_��Ժ����), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdDateInfo(DC_��Ժʱ��), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdDateInfo(DC_��Ժʱ��), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdInfo(GC_��Ժ����), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdInfo(GC_��Ժ����), Not blnEditInfo, , Not blnEditInfo)
        '������ҳ���ٱ༭����
        .lblNote.Visible = Not gclsPros.EditUnrecive
        .cboSpecificInfo(SLC_����).TabStop = gclsPros.TabAgeUnit
        .cboBaseInfo(BCC_����).TabStop = gclsPros.TabNation
        .txtInfo(GC_������).TabStop = gclsPros.TabFileNo
        .txtInfo(GC_������).TabStop = gclsPros.EditPageNo And gclsPros.OpenMode = EM_��������
        .txtInfo(GC_������).Locked = Not gclsPros.EditPageNo
        .cboBaseInfo(BCC_���ʽ).TabStop = gclsPros.TabPayType
        .chkInfo(CHK_����Ժ).TabStop = gclsPros.TabReadm
        .txtInfo(GC_X�ߺ�).TabStop = gclsPros.TabXRaysNo
        If Not gclsPros.EditPayType Then
            If .txtInfo(GC_������).Locked = False Then
                .txtInfo(GC_������).SetFocus
            ElseIf .txtSpecificInfo(SLC_סԺ��).Locked Then
                .txtInfo(GC_����).SetFocus
            Else
                .txtSpecificInfo(SLC_סԺ��).SetFocus
            End If
        Else
            '��Ҫ���ȱ���ҽ�Ƹ��ʽ
            .cboBaseInfo(BCC_���ʽ).SetFocus
        End If
        
        .vsOPS.ColHidden(PI_������ʿ) = Not gclsPros.Is����
    End With
End Sub

Public Sub LoadDataFromOutFile(ByVal strסԺ�� As String)
'�˴������Ϳ�ֵ������Ϊ�е�ϵͳ������û����Щֵ����
    Dim arrFileds As Variant, i As Long, strName As String

    On Error Resume Next
    arrFileds = Array("����", "��������", "������", "���֤��", "�Ա�", "Ѫ��", "ְҵ", "����", "����", "����״��", "��ϵ�˹�ϵ", "��λ�绰", "��λ�ʱ�", "��λ��ַ", _
                                "��ͥ��ַ", "���ڵ�ַ", "���ڵ�ַ�ʱ�", "��ͥ��ַ�ʱ�", "��ϵ�˵绰", "��ϵ�˵�ַ", "��ϵ������", "ҽ�Ƹ��ʽ", "��Ժ����", "����֤��", _
                                "��Ժ����", "��Ժ����", "סԺҽʦ", "���λ�ʿ")
    For i = LBound(arrFileds) To UBound(arrFileds)
        strName = IIf(arrFileds(i) = "������", "�����ص�", arrFileds(i))
        If Not IsNull(gclsPros.PatiOut(arrFileds(i))) Then
            Call SetCtrlValues(UCase(strName), gclsPros.PatiOut(arrFileds(i)) & "", , True)
        End If
    Next
    On Error GoTo errH
    '���˺�:20040812���ĵ�
    gclsPros.FeesOut.Filter = "סԺ�� = " & IIf(strסԺ�� = "", 0, strסԺ��) & " and סԺ����=" & gclsPros.��ҳID
    Call CacheLoadVsFreesData(gclsPros.CurrentForm.vsFees, gclsPros.FeesOut, , gclsPros.Is��Ŀ)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadDataFromPaitiInfo(ByVal lng����ID As Long, Optional ByVal intסԺ���� As Integer = 1) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id����ҳid ,��ȡ��صĲ�����Ϣ,����䵽��صĿؼ���
    '����:lng����ID-����id
    '     intסԺ����-��ҳID
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '�޸�:2007/08/29
    '----------------------------------------------------------------------------------------------------------------------------
    Dim rs������Ϣ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lng��ҳID As Long
    Dim strSql As String
    Dim i As Long
    Dim strCode As String

    lng��ҳID = IIf(gclsPros.SinPageNo, intסԺ����, 0)
    Set gclsPros.PatiInfo = GetPatiMainInfoData(lng����ID, lng��ҳID)
    '���ز�����Ϣ
    If Not gclsPros.PatiInfo.EOF Then
        For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
            If Not IsNull(gclsPros.PatiInfo.Fields(i).Value) Then
                Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
            End If
        Next
    End If
    Err = 0: On Error GoTo errH
    '������ҳסԺ�ţ������ţ������ŵȵ�����
    strCode = gclsPros.PatiInfo!��Ժ���ұ��� & ""
    If strCode = "" Then strCode = gclsPros.PatiInfo!�����ұ��� & ""
    'סԺ�Ż�ȡ
    If IsNull(gclsPros.PatiInfo!סԺ��) Then
        gclsPros.InNo = NVL(GetNextNo(2))
    ElseIf gclsPros.NewInNo And IsHavePageNos(CT_סԺ��, Not gclsPros.OpenMode = EM_�༭ Or gclsPros.Is��Ŀ, gclsPros.PatiInfo!סԺ�� & "", gclsPros.����ID) Then
        gclsPros.InNo = NVL(GetNextNo(2))
    Else
        gclsPros.InNo = gclsPros.PatiInfo!סԺ�� & ""
    End If
    gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = gclsPros.InNo
    '�����Ż�ȡ
    If IsNull(gclsPros.PatiInfo!������) Then
        If gclsPros.NewInNo Or Not gclsPros.SinPageNo And IsNull(gclsPros.PatiInfo!��󲡰���) Then
            '�����ʹ���µ�סԺ��,������ǿ��Ĭ��ΪסԺ��
            '���������סԺ������ , �򲡰��� = ��ǰסԺ��
            gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text
        ElseIf gclsPros.SinPageNo Then
            gclsPros.CurrentForm.txtInfo(GC_������).Text = NVL(GetNextNo(4, , strCode))
        ElseIf Not IsNull(gclsPros.PatiInfo!��󲡰���) Then
            '�����ǰ����������סԺ������,��ȡ���һ���������Ĳ�����
            gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.PatiInfo!��󲡰��� & ""
        End If
    Else
        gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.PatiInfo!������ & ""
    End If
    '53638:������,2013-05-10,���������ű�Ź���
    If IsNull(gclsPros.PatiInfo!��󵵰���) And gclsPros.UseFileRules Then
        gclsPros.CurrentForm.txtInfo(GC_������).Text = NVL(GetNextNo(5, , strCode))
    Else
        gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.PatiInfo!��󵵰��� & ""
    End If
    Call CacheLoadVsAllerData(gclsPros.CurrentForm.vsAller, GetAllerData(lng����ID, lng��ҳID))
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function SetCboDefaultValue(ByVal lngIndex As Long) As Boolean
'----------------------------------------------------------------------------------------------------------------------------
'����:�����ֵ����������ĳһComboBox ��ֵ��Ĭ��ֵ
'����:lngIndex: cboBaseInfo �ؼ�������
'����:�ɹ�,����true,���򷵻�False
'----------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim rsTmp As ADODB.Recordset
    Dim objCboTmp As ComboBox
    On Error GoTo errH
    Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(lngIndex)
    Set rsTmp = GetBaseCode(lngIndex)
    '���ԭ������
    objCboTmp.Clear
    objCboTmp.Tag = ""
    'װ������
    If Not rsTmp.EOF Then
        For j = 1 To rsTmp.RecordCount
            If IsNull(rsTmp!����) Then
                objCboTmp.AddItem rsTmp!����
            Else
                objCboTmp.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����
            End If
            objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
            If Val(rsTmp!ȱʡ & "") = 1 Then
                Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                objCboTmp.Tag = objCboTmp.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    SetCboDefaultValue = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetMainDirectory()
'����:����ҳ�����õ���Ŀ¼
    Dim myNod As Node
    Dim strTmp As String
    Dim strTitle() As String
    Dim i As Long, j As Long
    
    If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
        strTmp = "����,������Ϣ,��ҽ���,��ҽ������,��ҽ���,��ҽ������,ҩ�����,��Ѫ��Ϣ,ǩ����Ϣ,������¼,סԺ����,סԺ���,������Ϣ,������Ϣ,����ҩ��ʹ�����,�������������,��֢�໤���,����������,��ҳ1,��ҳ2"
    Else
        strTmp = "����,������Ϣ,��ҽ���,��ҽ������,��ҽ���,��ҽ������,ҩ�����,��Ѫ��Ϣ,ǩ����Ϣ,������¼,סԺ����,סԺ���,������Ϣ,������Ϣ,����ҩ��ʹ�����,�������������,��֢�໤���,����������,��ҳ"
    End If
   
    '������Ҹ�ҳĿ¼
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            strTmp = strTmp & "," & gfrmMecCol(i).Caption
        Next
    End If
    
    j = 1
    strTitle = Split(strTmp, ",")
    
    frmMain.tvDirectory.Nodes.Clear
    frmMain.tvDirectory.LineStyle = tvwRootLines
    frmMain.tvDirectory.Indentation = 200
    
    With gclsPros.CurrentForm
        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                Set myNod = frmMain.tvDirectory.Nodes.Add(, , "key-" & i, j & ". " & strTitle(i))
                myNod.Expanded = True
                j = j + 1
            End If
        Next
    End With
    
End Sub

Public Function GetReplaceObject(ByVal vsfTmp As VSFlexGrid) As TextBox
'����: ���ݴ����VSFlexGrid�ڵ�������г�����һ�����ص�TextBox�ؼ�
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim i As Long
    Dim vPoint As POINTAPI
    
    With gclsPros.CurrentForm
        If .txtInfo.UBound <> 9999 Then
            Load .txtInfo(9999)
        End If
        vPoint = GetCoordPos(vsfTmp.hwnd, vsfTmp.CellLeft, vsfTmp.CellTop)
        .txtInfo(9999).Visible = False
        Set .txtInfo(9999).Container = vsfTmp.Container
        lngLeft = vPoint.X - frmMain.Left - .picMain.Left - vsfTmp.Container.Left - frmMain.PicDirectory.Width - 200
        lngTop = vPoint.Y - frmMain.Top - frmMain.PicForm.Top - .picMain.Top - vsfTmp.Container.Top - .Top - 80
        .txtInfo(9999).Move lngLeft, lngTop, vsfTmp.ColWidth(vsfTmp.Col), vsfTmp.RowHeight(vsfTmp.Row)

        Set GetReplaceObject = .txtInfo(9999)
    End With
End Function


Private Sub LoadDiagAndAllerFData()
'���ܣ���������֮�����¼�����ϣ�������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long
    Dim intMaxDiagSource As Integer
    Dim LngRow As Long, vsTmp As VSFlexGrid
    
    With gclsPros.CurrentForm
        lng����ID = gclsPros.����ID
        lng��ҳID = gclsPros.��ҳID
        
        '������Ϣ����
        If gclsPros.AddAller Then
            Call DeleteCacheRecInfo("����ҩ��")
            Set rsTmp = GetAllerData(lng����ID, lng��ҳID)
            Call CacheLoadVsAllerData(.vsAller, rsTmp)
            gclsPros.AddAller = False
        End If
    
        '��ȡ���
        If gclsPros.AddDiag Then
            Set rsTmp = GetPatiDiagData(lng����ID, lng��ҳID, IIf(gclsPros.PatiType <> PF_����, 1, 0), , Not gclsPros.Is��Ŀ, gclsPros.Moved)
            rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
            intMaxDiagSource = IIf(gclsPros.FuncType = f������ҳ, 4, -1)
            If gclsPros.FuncType = f������ҳ And rsTmp.EOF Then
                intMaxDiagSource = 3
                rsTmp.Filter = "��¼��Դ=3"
                If rsTmp.EOF Then intMaxDiagSource = 2
            End If
            If Not gclsPros.Is���� Or gclsPros.Is���� And rsTmp.RecordCount = 0 Then
                '2��������ҽ���
                Call DeleteCacheRecInfo("��ҽ���")
                Call InitTableDiag
                Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "1,2,3,5,6,7,10", "1"), , intMaxDiagSource)
                '3��������ҽ���
                If gclsPros.Have��ҽ Then
                    Call DeleteCacheRecInfo("��ҽ���")
                    Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "11,12,13", "11"), , intMaxDiagSource)
                End If
                gclsPros.AddDiag = False
            End If
            
            Set vsTmp = .vsDiagXY
            With vsTmp
                .Cell(flexcpForeColor, 1, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
                .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
                If gclsPros.PatiType <> PF_���� Then
                    LngRow = FindDiagRow(DT_��Ժ���XY)
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                    .Row = .FixedRows: .Col = DI_�������
                    Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
                Else
                    .Cell(flexcpText, .FixedRows, DI_�������, .Rows - 1, DI_�������) = "��ҽ"
                End If
            End With
    
            Set vsTmp = .vsDiagZY
            With vsTmp
                .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
                .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
                If gclsPros.PatiType <> PF_���� Then
                    LngRow = FindDiagRow(DT_��Ժ���ZY)
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                    Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
                Else
                    .Cell(flexcpText, .FixedRows, DI_�������, .Rows - 1, DI_�������) = "��ҽ"
                End If
            End With
            
        End If
    End With
End Sub

Public Sub DeleteCacheRecInfo(ByVal strInfoName As String)
'���ܣ�ɾ����Ϣ��¼����һ��Ӧ���ڱ��
'������strInfoName=��Ϣ����ؼ���
    On Error GoTo errH
    '��������Ϣ��Ѱ��Ѱ�ң�Ѱ�Ҳ���ʱ���ٰ��ؼ���Ѱ��
    gclsPros.MainInfoRec.Filter = "��Ϣ��='" & strInfoName & "'"
    If gclsPros.MainInfoRec.EOF Then gclsPros.MainInfoRec.Filter = "�ؼ���='" & strInfoName & "'"
    If Not gclsPros.MainInfoRec.EOF Then
        Select Case gclsPros.MainInfoRec!ExpState
            Case ES_������չ
                Call Rec.Delete(gclsPros.SecdInfoRec, "���=" & gclsPros.MainInfoRec!���)
        End Select
    End If
    Exit Sub
errH:
    Debug.Print "DeleteCacheRecInfo:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckValueChange(Optional ByRef objTmp As Object) As Boolean
'���ܣ������ҳ�ؼ���ֵ�Ƿ����仯
    Dim strOlsInfo As String
    Dim strCurInfo As String
    Dim strCboName As String
    Dim cboTmp As ComboBox
    Dim lngIndex As Long
    Dim blnFind As Boolean
    
    If gclsPros.InfosChange Then Exit Function
    If Not gclsPros.LoadFinish Then Exit Function
    If gclsPros.FuncType <> fҽ����ҳ And gclsPros.FuncType <> f������ҳ Then Exit Function
    If gclsPros.OpenMode = EM_���� Then Exit Function
    On Error GoTo errH
    If frmMain.stbThis.Panels(2).Text <> "" Then
        frmMain.stbThis.Panels(2).Text = ""
    End If
    If objTmp Is Nothing Then
        gclsPros.InfosChange = True
        Exit Function
    End If
    If TypeName(objTmp) = "ComboBox" Then
        Set cboTmp = objTmp
        strCurInfo = cboTmp.Text
        strCboName = cboTmp.Name
        lngIndex = cboTmp.Index
    Else
        gclsPros.InfosChange = True
        Exit Function
    End If
    
    If strCboName = "cboBaseInfo" Or strCboName = "cboManInfo" Then
        gclsPros.MainInfoRec.Filter = "�ؼ���='" & strCboName & "'" & "And Index=" & lngIndex
        If Not gclsPros.MainInfoRec.EOF Then
            strOlsInfo = NVL(gclsPros.MainInfoRec!��Ϣԭֵ)
            blnFind = True
        Else
            gclsPros.SecdInfoRec.Filter = "�ؼ���='" & strCboName & "'" & "And IndexEx=" & lngIndex
            If Not gclsPros.SecdInfoRec.EOF Then
                strOlsInfo = NVL(gclsPros.SecdInfoRec!��Ϣԭֵ)
                blnFind = True
            End If
        End If
        If blnFind Then
            If strCurInfo <> strOlsInfo And blnFind Then
                gclsPros.InfosChange = True
            End If
        Else
            gclsPros.InfosChange = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LocateVSFRowCol(ByRef vsfTmp As VSFlexGrid, ByVal lngMinRow As Long, ByVal lngMaxRow As Long, ByVal lngMinCol As Long, ByVal lngMaxCol As Long, ByVal LngRow As Long, ByVal LngCol As Long)
    If Not Between(vsfTmp.Row, lngMinRow, lngMaxRow) Then vsfTmp.Row = LngRow
    If Not Between(vsfTmp.Col, lngMinCol, lngMaxCol) Then vsfTmp.Col = LngCol
    If vsfTmp.ColHidden(vsfTmp.Col) = True Then vsfTmp.Col = LngCol
End Sub

Private Sub SetDeliceryInfo(ByRef vsDiagTmp As VSFlexGrid)
'���ܣ�������������Ϣ�����á�������Ϣ����ť�Ŀɼ���
    Dim bln��ҽ As Boolean
    Dim lngTmpRow As Long, i As Long, j As Long

    On Error GoTo errH
    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        If gclsPros.PatiType <> PF_���� Then
            If gclsPros.FuncType = f������ҳ Then
                If bln��ҽ Then                             '��������
                    lngTmpRow = FindDiagRow(DT_�������)
                    i = FindDiagRow(DT_��Ժ���XY)
                    
                    gclsPros.CurrentForm.cmdDeliceryInfo.Visible = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Tag = ""
                    
                    For j = i To lngTmpRow - 1
                        If .TextMatrix(j, DI_������Ϣ) = "1" Then
                            gclsPros.CurrentForm.cmdDeliceryInfo.Visible = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'�����ַ������������ʾ
Public Sub ErrDw(strMsg As String)
    'strMsg���� ��: ��ʾ��Ϣ-(����:1/��ֹ:0)|�ؼ�Keyֵ-��ʾ��Ϣ-(����:1/��ֹ:0)|���ؼ�Keyֵ-��ʾ��Ϣ-(����:1/��ֹ:0)-Row-Col
    Dim i As Long
    Dim arrTmp() As String
    
    On Error GoTo errH
    With gclsPros.CurrentForm
        If strMsg <> "" Then
            ReDim Preserve arrTmp(UBound(Split(strMsg, "|")))
            arrTmp = Split(strMsg, "|")
            For i = 0 To UBound(arrTmp)
                Select Case UBound(Split(arrTmp(i), "-"))
                    Case 1 'ֻ��ʾ��Ϣ���󶨿ؼ�
                         Call AddErrInfo(Split(arrTmp(i), "-")(0), Val(Split(arrTmp(i), "-")(1)))
                    Case 2 '�󶨿ؼ���ʾ��Ϣ
                         Call AddErrInfo(Split(arrTmp(i), "-")(1), Val(Split(arrTmp(i), "-")(2)), gColCtl.Item((Split(arrTmp(i), "-")(0))))
                    Case 4 '�󶨱��ؼ���ʾ��Ϣ
                        gColCtl.Item((Split(arrTmp(i), "-")(0))).Row = Val((Split(arrTmp(i), "-")(3)))
                        gColCtl.Item((Split(arrTmp(i), "-")(0))).Col = Val((Split(arrTmp(i), "-")(4)))
                        Call AddErrInfo(Split(arrTmp(i), "-")(1), Val(Split(arrTmp(i), "-")(2)), gColCtl.Item((Split(arrTmp(i), "-")(0))))
                End Select
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'����������Ϣ���������ʾ
Public Sub ErrMec(colMsg As Collection)
    'colMsg���� ��: �ؼ�����:colMsg.item,colMsg.item.tag((����:1/��ֹ:0)|��ʾ��Ϣ|PicPage��Index)
    Dim i As Long
    Dim arrTmp() As String
    
    On Error GoTo errH
    With gclsPros.CurrentForm
        For i = 1 To colMsg.Count
            ReDim Preserve arrTmp(UBound(Split(colMsg.Item(i).Tag, "|")))
            arrTmp = Split(colMsg.Item(i).Tag, "|")
            If UBound(arrTmp) <> 0 Then
                Call AddErrInfo(arrTmp(1), Val(arrTmp(0)), colMsg.Item(i))
            End If
            colMsg.Item(i).Tag = ""
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function CtlAdd() As Collection
    Dim colCtl As New Collection
    Dim i As Long
    On Error Resume Next
    With gclsPros.CurrentForm
        colCtl.Add .cboManInfo(MC_����ҽʦ), "����ҽʦ"
        colCtl.Add .cboManInfo(MC_������), "������"
        colCtl.Add .cboManInfo(MC_���λ�����), "���λ�����"
        colCtl.Add .cboManInfo(MC_����ҽʦ), "����ҽʦ"
        colCtl.Add .cboManInfo(MC_����ҽʦ), "����ҽʦ"
        colCtl.Add .cboManInfo(MC_סԺҽʦ), "סԺҽʦ"
        colCtl.Add .cboManInfo(MC_�о���ҽʦ), "�о���ҽʦ"
        colCtl.Add .cboManInfo(MC_ʵϰҽʦ), "ʵϰҽʦ"
        colCtl.Add .cboManInfo(MC_�ʿ�ҽʦ), "�ʿ�ҽʦ"
        colCtl.Add .cboManInfo(MC_�ʿػ�ʿ), "�ʿػ�ʿ"
        colCtl.Add .cboManInfo(MC_���λ�ʿ), "���λ�ʿ"
        colCtl.Add .cboManInfo(MC_��ĿԱ), "��ĿԱ"
        colCtl.Add .cboManInfo(MC_����ҽʦ), "����ҽʦ"
        colCtl.Add .mskDateInfo(DC_��������), "��������"
        colCtl.Add .mskDateInfo(DC_��Ժʱ��), "��Ժʱ��"
        colCtl.Add .mskDateInfo(DC_��Ժʱ��), "��Ժʱ��"
        colCtl.Add .mskDateInfo(DC_ȷ������), "ȷ������"
        colCtl.Add .mskDateInfo(DC_����ʱ��), "����ʱ��"
        colCtl.Add .mskDateInfo(DC_��������), "��������"
        colCtl.Add .mskDateInfo(DC_����ʱ��), "����ʱ��"
        colCtl.Add .mskDateInfo(DC_�ʿ�����), "�ʿ�����"
        colCtl.Add .mskDateInfo(DC_��Ŀ����), "��Ŀ����"
        colCtl.Add .mskDateInfo(DC_�ջ�����), "�ջ�����"
        colCtl.Add .txtAdressInfo(ADRC_�����ص�), "�����ص�"
        colCtl.Add .txtAdressInfo(ADRC_����), "����"
        colCtl.Add .txtAdressInfo(ADRC_��סַ), "��סַ"
        colCtl.Add .txtAdressInfo(ADRC_���ڵ�ַ), "���ڵ�ַ"
        colCtl.Add .txtAdressInfo(ADRC_��ϵ�˵�ַ), "��ϵ�˵�ַ"
        colCtl.Add .txtAdressInfo(ADRC_��������), "��������"
        colCtl.Add .txtAdressInfo(ADRC_��λ��ַ), "��λ��ַ"
        colCtl.Add .padrInfo(ADRC_�����ص�), "�����ص�(�ṹ��)"
        colCtl.Add .padrInfo(ADRC_����), "����(�ṹ��)"
        colCtl.Add .padrInfo(ADRC_��סַ), "��סַ(�ṹ��)"
        colCtl.Add .padrInfo(ADRC_���ڵ�ַ), "���ڵ�ַ(�ṹ��)"
        colCtl.Add .padrInfo(ADRC_��ϵ�˵�ַ), "��ϵ�˵�ַ(�ṹ��)"
        colCtl.Add .cboBaseInfo(BCC_���ʽ), "���ʽ"
        colCtl.Add .cboBaseInfo(BCC_�Ա�), "�Ա�"
        colCtl.Add .cboBaseInfo(BCC_����), "����"
        colCtl.Add .cboBaseInfo(BCC_ְҵ), "ְҵ"
        colCtl.Add .cboBaseInfo(BCC_����), "����"
        colCtl.Add .cboBaseInfo(BCC_����), "����"
        colCtl.Add .cboBaseInfo(BCC_��ϵ), "��ϵ"
        colCtl.Add .cboBaseInfo(BCC_��Ժ;��), "��Ժ;��"
        colCtl.Add .cboBaseInfo(BCC_�Ļ��̶�), "�Ļ��̶�"
        colCtl.Add .cboBaseInfo(BCC_ȥ��), "ȥ��"
        colCtl.Add .cboBaseInfo(BCC_��Ⱦ��������ϵ), "��Ⱦ��������ϵ"
        colCtl.Add .cboBaseInfo(BCC_��Ժ���), "��Ժ���"
        colCtl.Add .cboBaseInfo(BCC_�ֻ��̶�), "�ֻ��̶�"
        colCtl.Add .cboBaseInfo(BCC_����������), "����������"
        colCtl.Add .cboBaseInfo(BCC_�������ԺXY), "�������ԺXY"
        colCtl.Add .cboBaseInfo(BCC_��Ժ���ԺXY), "��Ժ���ԺXY"
        colCtl.Add .cboBaseInfo(BCC_��������Ժ), "��������Ժ"
        colCtl.Add .cboBaseInfo(BCC_��ǰ������), "��ǰ������"
        colCtl.Add .cboBaseInfo(BCC_�����벡��), "�����벡��"
        colCtl.Add .cboBaseInfo(BCC_�ٴ��벡��), "�ٴ��벡��"
        colCtl.Add .cboBaseInfo(BCC_�����ڼ�), "�����ڼ�"
        colCtl.Add .cboBaseInfo(BCC_�ٴ���ʬ��), "�ٴ���ʬ��"
        colCtl.Add .cboBaseInfo(BCC_�������ԺZY), "�������ԺZY"
        colCtl.Add .cboBaseInfo(BCC_��Ժ���ԺZY), "��Ժ���ԺZY"
        colCtl.Add .cboBaseInfo(BCC_��֤), "��֤"
        colCtl.Add .cboBaseInfo(BCC_�η�), "�η�"
        colCtl.Add .cboBaseInfo(BCC_��ҩ), "��ҩ"
        colCtl.Add .cboBaseInfo(BCC_�������), "�������"
        colCtl.Add .cboBaseInfo(BCC_��ҽ�����豸), "��ҽ�����豸"
        colCtl.Add .cboBaseInfo(BCC_���ȷ���), "���ȷ���"
        colCtl.Add .cboBaseInfo(BCC_��ҽ���Ƽ���), "��ҽ���Ƽ���"
        colCtl.Add .cboBaseInfo(BCC_������ҩ�Ƽ�), "������ҩ�Ƽ�"
        colCtl.Add .cboBaseInfo(BCC_��֤ʩ��), "��֤ʩ��"
        colCtl.Add .cboBaseInfo(BCC_��������), "��������"
        colCtl.Add .cboBaseInfo(BCC_��������), "��������"
        colCtl.Add .cboBaseInfo(BCC_HBsAg), "HBsAg"
        colCtl.Add .cboBaseInfo(BCC_Ѫ��), "Ѫ��"
        colCtl.Add .cboBaseInfo(BCC_HCVAb), "HCVAb"
        colCtl.Add .cboBaseInfo(BCC_RH), "RH"
        colCtl.Add .cboBaseInfo(BCC_HIVAb), "HIVAb"
        colCtl.Add .cboBaseInfo(BCC_��Һ��Ӧ), "��Һ��Ӧ"
        colCtl.Add .cboBaseInfo(BCC_��Ѫ��Ӧ), "��Ѫ��Ӧ"
        colCtl.Add .cboBaseInfo(BCC_��Ѫǰ9����), "��Ѫǰ9����"
        colCtl.Add .cboBaseInfo(BCC_����״��), "����״��"
        colCtl.Add .cboBaseInfo(BCC_��Ժ��ʽ), "��Ժ��ʽ"
        colCtl.Add .cboBaseInfo(BCC_����Ժ�ƻ�����), "����Ժ�ƻ�����"
        colCtl.Add .cboBaseInfo(BCC_ѹ�������ڼ�), "ѹ�������ڼ�"
        colCtl.Add .cboBaseInfo(BCC_ѹ������), "ѹ������"
        colCtl.Add .cboBaseInfo(BCC_������׹���˺�), "������׹���˺�"
        colCtl.Add .cboBaseInfo(BCC_������׹��ԭ��), "������׹��ԭ��"
        colCtl.Add .cboBaseInfo(BCC_���ϴ�סԺʱ��), "���ϴ�סԺʱ��(����)"
        colCtl.Add .cboBaseInfo(BCC_�ط����ʱ��), "�ط����ʱ��"
        colCtl.Add .cboBaseInfo(BCC_Լ����ʽ), "Լ����ʽ"
        colCtl.Add .cboBaseInfo(BCC_Լ������), "Լ������"
        colCtl.Add .cboBaseInfo(BCC_Լ��ԭ��), "Լ��ԭ��"
        colCtl.Add .cboBaseInfo(BCC_��������Ժ��ʽ), "��������Ժ��ʽ"
        colCtl.Add .cboBaseInfo(BCC_��������), "��������"
        colCtl.Add .cboBaseInfo(BCC_�ٴ�·������), "�ٴ�·������"
        colCtl.Add .cboBaseInfo(BCC_������Ⱦ��), "������Ⱦ��"
        colCtl.Add .cboBaseInfo(BCC_ʵʩDGRS����), "ʵʩDGRS����"
        colCtl.Add .cboBaseInfo(BCC_��������ʬ��), "��������ʬ��"
        colCtl.Add .cboBaseInfo(BCC_���֤), "���֤"
        colCtl.Add .cboBaseInfo(BCC_����ԭ��), "����ԭ��(����)"
        colCtl.Add .cboBaseInfo(BCC_��������), "��������"
        colCtl.Add .chkInfo(CHK_����Ժ), "����Ժ"
        colCtl.Add .chkInfo(CHK_��Ժǰ��Ժ����), "��Ժǰ��Ժ����"
        colCtl.Add .chkInfo(CHK_�Ƿ�ȷ��), "�Ƿ�ȷ��"
        colCtl.Add .chkInfo(CHK_��ԭѧ���), "��ԭѧ���"
        colCtl.Add .chkInfo(CHK_�·�����), "�·�����"
        colCtl.Add .chkInfo(CHK_Σ��), "Σ��"
        colCtl.Add .chkInfo(CHK_��֢), "��֢"
        colCtl.Add .chkInfo(CHK_����), "����"
        colCtl.Add .chkInfo(CHK_ʾ�̲���), "ʾ�̲���"
        colCtl.Add .chkInfo(CHK_���в���), "���в���"
        colCtl.Add .chkInfo(CHK_���Ѳ���), "���Ѳ���"
        colCtl.Add .chkInfo(CHK_����), "����"
        colCtl.Add .chkInfo(CHK_CT), "CT"
        colCtl.Add .chkInfo(CHK_MRI), "MRI"
        colCtl.Add .chkInfo(CHK_��ɫ������), "��ɫ������"
        colCtl.Add .chkInfo(CHK_��Ⱦ���ϴ�), "��Ⱦ���ϴ�"
        colCtl.Add .chkInfo(CHK_Χ��������), "Χ��������"
        colCtl.Add .chkInfo(CHK_�������), "�������"
        colCtl.Add .chkInfo(CHK_����·��), "����·��"
        colCtl.Add .chkInfo(CHK_���·��), "���·��"
        colCtl.Add .chkInfo(CHK_����), "����"
        colCtl.Add .chkInfo(CHK_סԺ����Σ��), "סԺ����Σ��"
        colCtl.Add .chkInfo(CHK_�Ƿ�ͬһ����), "�Ƿ�ͬһ����"
        colCtl.Add .chkInfo(CHK_�˹������ѳ�), "�˹������ѳ�"
        colCtl.Add .chkInfo(CHK_�ط���֢ҽѧ��), "�ط���֢ҽѧ��"
        colCtl.Add .chkInfo(CHK_סԺ����Լ��), "סԺ����Լ��"
        colCtl.Add .chkInfo(CHK_�����ֹ���), "�����ֹ���"
        colCtl.Add .chkInfo(CHK_ϸ���걾�ͼ�), "ϸ���걾�ͼ�"
        colCtl.Add .chkInfo(CHK_�������), "�������"
        colCtl.Add .chkInfo(CHK_�޹�����¼), "�޹�����¼"
        colCtl.Add .txtInfo(GC_������), "������"
        colCtl.Add .txtInfo(GC_������), "������"
        colCtl.Add .txtInfo(GC_X�ߺ�), "X�ߺ�"
        colCtl.Add .txtInfo(GC_����), "����"
        colCtl.Add .txtInfo(GC_����֤��), "����֤��"
        colCtl.Add .txtInfo(GC_��ϵ������), "��ϵ������"
        colCtl.Add .txtInfo(GC_��Ժ����), "��Ժ����"
        colCtl.Add .txtInfo(GC_��Ժ����), "��Ժ����"
        colCtl.Add .txtInfo(GC_��Ժ����), "��Ժ����"
        colCtl.Add .txtInfo(GC_��Ժ����), "��Ժ����"
        colCtl.Add .txtInfo(GC_ҽ����), "ҽ����"
        colCtl.Add .txtInfo(GC_ժҪ), "ժҪ"
        colCtl.Add .txtInfo(GC_�����), "�����"
        colCtl.Add .txtInfo(GC_�໤��), "�໤��"
        colCtl.Add .txtInfo(GC_������ַ), "������ַ"
        colCtl.Add .txtInfo(GC_ҽѧ��ʾ), "ҽѧ��ʾ"
        colCtl.Add .txtInfo(GC_����ҽѧ��ʾ), "����ҽѧ��ʾ"
        colCtl.Add .txtInfo(GC_�����), "�����"
        colCtl.Add .txtInfo(GC_����ԭ��), "����ԭ��"
        colCtl.Add .txtInfo(GC_��ԭѧ���), "��ԭѧ���"
        colCtl.Add .txtInfo(GC_���Ȳ���), "���Ȳ���"
        colCtl.Add .txtInfo(GC_������), "������"
        colCtl.Add .txtInfo(GC_ת��ҽ�ƻ���), "ת��ҽ�ƻ���"
        colCtl.Add .txtInfo(GC_31������סԺ), "31������סԺ"
        colCtl.Add .txtInfo(GC_ת��1), "ת��1"
        colCtl.Add .txtInfo(GC_ת��2), "ת��2"
        colCtl.Add .txtInfo(GC_ת��3), "ת��3"
        colCtl.Add .txtInfo(GC_�˳�ԭ��), "�˳�ԭ��"
        colCtl.Add .txtInfo(GC_����ԭ��), "����ԭ��"
        colCtl.Add .txtInfo(GC_��֢�໤������), "��֢�໤������"
        colCtl.Add .txtInfo(GC_����T), "����T"
        colCtl.Add .txtInfo(GC_����N), "����N"
        colCtl.Add .txtInfo(GC_����M), "����M"
        colCtl.Add .txtInfo(GC_Email), "Email"
        colCtl.Add .txtInfo(GC_��������), "��������"
        colCtl.Add .txtInfo(GC_����ҩ��), "����ҩ��"
        colCtl.Add .txtInfo(GC_�ٴ�����), "�ٴ�����"
        colCtl.Add .txtInfo(GC_͸�����ص�ֵ), "͸�����ص�ֵ"
        colCtl.Add .txtInfo(GC_������ϵ), "������ϵ"
        colCtl.Add .txtInfo(GC_��Ժת��), "��Ժת��"
        colCtl.Add .optInput(OP_��סԺ��), "��סԺ��"
        colCtl.Add .optInput(OP_��סԺ��), "��סԺ��"
        colCtl.Add .optInput(OP_����), "����"
        colCtl.Add .optInput(OP_����), "����"
        colCtl.Add .optInput(OP_ICU��), "ICU��"
        colCtl.Add .optInput(OP_ICU��), "ICU��"
        colCtl.Add .optDiag(PC_XY���������), "XY���������"
        colCtl.Add .optDiag(PC_XY��������������), "XY��������������"
        colCtl.Add .optDiag(PC_ZY���������), "ZY���������"
        colCtl.Add .optDiag(PC_ZY��������������), "ZY��������������"
        colCtl.Add .optDiag(PC_���������), "���������"
        colCtl.Add .optDiag(PC_��������������), "��������������"
        colCtl.Add .optAller(PC_��ҩƷĿ¼����), "��ҩƷĿ¼����"
        colCtl.Add .optAller(PC_������Դ����), "������Դ����"
        colCtl.Add .OptParaOPSInfo(PC_��������Ŀ����), "��������Ŀ����"
        colCtl.Add .OptParaOPSInfo(PC_��ICDCM9��������), "��ICDCM9��������"
        colCtl.Add .chkParaOPSInfo(PC_δ�ҵ�ʱ����¼��), "δ�ҵ�ʱ����¼��"
        colCtl.Add .cboSpecificInfo(SLC_����), "����(����)"
        colCtl.Add .cboSpecificInfo(SLC_Ӥ�׶�����), "Ӥ�׶�����(����)"
        colCtl.Add .cboSpecificInfo(SLC_��������), "��������(����)"
        colCtl.Add .txtSpecificInfo(SLC_��λ�绰), "��λ�绰"
        colCtl.Add .txtSpecificInfo(SLC_��λ�ʱ�), "��λ�ʱ�"
        colCtl.Add .txtSpecificInfo(SLC_��ͥ�绰), "��ͥ�绰"
        colCtl.Add .txtSpecificInfo(SLC_��ͥ�ʱ�), "��ͥ�ʱ�"
        colCtl.Add .txtSpecificInfo(SLC_�����ʱ�), "�����ʱ�"
        colCtl.Add .txtSpecificInfo(SLC_���), "���"
        colCtl.Add .txtSpecificInfo(SLC_��ߵ�λ), "��ߵ�λ"
        colCtl.Add .txtSpecificInfo(SLC_����), "����"
        colCtl.Add .txtSpecificInfo(SLC_���ص�λ), "���ص�λ"
        colCtl.Add .txtSpecificInfo(SLC_����), "����"
        colCtl.Add .txtSpecificInfo(SLC_��Ժ����), "��Ժ����"
        colCtl.Add .txtSpecificInfo(SLC_����ѹ), "����ѹ"
        colCtl.Add .txtSpecificInfo(SLC_����ѹ), "����ѹ"
        colCtl.Add .txtSpecificInfo(SLC_��ϵ�˵绰), "��ϵ�˵绰"
        colCtl.Add .txtSpecificInfo(SLC_����), "����"
        colCtl.Add .txtSpecificInfo(SLC_Ӥ�׶�����), "Ӥ�׶�����"
        colCtl.Add .txtSpecificInfo(SLC_��������������), "��������������"
        colCtl.Add .txtSpecificInfo(SLC_��������Ժ����), "��������Ժ����"
        colCtl.Add .txtSpecificInfo(SLC_סԺ����), "סԺ����"
        colCtl.Add .txtSpecificInfo(SLC_סԺ��), "סԺ��"
        colCtl.Add .txtSpecificInfo(SLC_���ȴ���), "���ȴ���"
        colCtl.Add .txtSpecificInfo(SLC_�ɹ�����), "�ɹ�����"
        colCtl.Add .txtSpecificInfo(SLC_�ػ�), "�ػ�"
        colCtl.Add .txtSpecificInfo(SLC_һ������), "һ������"
        colCtl.Add .txtSpecificInfo(SLC_��������), "��������"
        colCtl.Add .txtSpecificInfo(SLC_��������), "��������"
        colCtl.Add .txtSpecificInfo(SLC_ICU), "ICU"
        colCtl.Add .txtSpecificInfo(SLC_CCU), "CCU"
        colCtl.Add .txtSpecificInfo(SLC_���ϸ��), "���ϸ��"
        colCtl.Add .txtSpecificInfo(SLC_��ѪС��), "��ѪС��"
        colCtl.Add .txtSpecificInfo(SLC_��Ѫ��), "��Ѫ��"
        colCtl.Add .txtSpecificInfo(SLC_��ȫѪ), "��ȫѪ"
        colCtl.Add .txtSpecificInfo(SLC_�������), "�������"
        colCtl.Add .txtSpecificInfo(SLC_������ʹ��), "������ʹ��"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժǰ_��), "����ʱ����Ժǰ_��"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժǰ_Сʱ), "����ʱ����Ժǰ_Сʱ"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժǰ_����), "����ʱ����Ժǰ_����"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժ��_��), "����ʱ����Ժ��_��"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժ��_Сʱ), "����ʱ����Ժ��_Сʱ"
        colCtl.Add .txtSpecificInfo(SLC_����ʱ����Ժ��_����), "����ʱ����Ժ��_����"
        colCtl.Add .txtSpecificInfo(SLC_��������), "��������"
        colCtl.Add .txtSpecificInfo(SLC_���ú�), "���ú�"
        colCtl.Add .txtSpecificInfo(SLC_Լ����ʱ��), "Լ����ʱ��"
        colCtl.Add .txtSpecificInfo(SLC_��֢�໤��), "��֢�໤��"
        colCtl.Add .txtSpecificInfo(SLC_��֢�໤Сʱ), "��֢�໤Сʱ"
        colCtl.Add .txtSpecificInfo(SLC_Apgar), "Apgar"
        colCtl.Add .txtSpecificInfo(SLC_QQ), "QQ"
        colCtl.Add .txtSpecificInfo(SLC_��׵���), "��׵���"
        colCtl.Add .txtSpecificInfo(SLC_Ժ�ڻ���), "Ժ�ڻ���"
        colCtl.Add .txtSpecificInfo(SLC_��Ժ����), "��Ժ����"
        colCtl.Add .txtSpecificInfo(SLC_���ϴ�סԺʱ��), "���ϴ�סԺʱ��"
        colCtl.Add .vsTransfer, "ת�����(���)"
        colCtl.Add .vsDiagXY, "��ҽ���(���)"
        colCtl.Add .vsDiagZY, "��ҽ���(���)"
        colCtl.Add .vsAller, "������Ϣ(���)"
        colCtl.Add .vsOPS, "������¼(���)"
        colCtl.Add .vsFees, "סԺ����(���)"
        colCtl.Add .vsChemoth, "���Ƽ�¼��Ϣ(���)"
        colCtl.Add .vsRadioth, "���Ƽ�¼��Ϣ(���)"
        colCtl.Add .vsKSS, "����ҩ��ʹ�����(���)"
        colCtl.Add .vsSpirit, "�������������(���)"
        colCtl.Add .vsFlxAddICU, "��֢�໤���(���)"
        colCtl.Add .vsfMain, "����������Ŀ(���)"
        colCtl.Add .vsTSJC, "���������(���)"
        colCtl.Add .lstAdvEvent, "�����¼�(���)"
        colCtl.Add .lstInfection, "��Ⱦ����(���)"
        colCtl.Add .lvwFee, "סԺ����(���α�)"
        colCtl.Add .padrInfo(ADRC_��λ��ַ), "��λ��ַ(�ṹ��)"
        colCtl.Add .txtSpecificInfo(SLC_Ӥ�׶�����_DAY), "Ӥ�׶�����_DAY"
        colCtl.Add .vsICUInstruments, "��е����ʹ�����(���)"
        colCtl.Add .vsInfect, "���˸�Ⱦ��¼(���)"
        colCtl.Add .lstInfectParts, "��Ⱦ��λ(���)"
        colCtl.Add .vsSample, "���˲�ԭѧ���(���)"
    End With
    Set CtlAdd = colCtl
    On Error GoTo 0
End Function

Private Sub MsgDis(str����IDs As String, str���IDs As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSql As String
    On Error GoTo Errhand
    '�жϵ�ǰ�����Ƿ���д��Ⱦ�����濨
    strSql = "Select �ļ�ID From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And ��������=5 and ������=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "MsgDis", gclsPros.����ID, gclsPros.��ҳID, UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        '�ж��û��Ƿ��޸Ļ�ɾ�����
        strSql = ""
        If str����IDs <> "" Then
            strSql = " Union Select ����id,���id From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
        End If
        If str���IDs <> "" Then
            strSql = strSql & " Union Select ����id,���id From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
        End If
        strSql = "Select a.����id, a.���id From ������ϼ�¼ A, ��������ǰ�� B Where a.����id = [1] And a.��ҳid = [2] And a.������� = 1 And (a.����id = b.����id Or a.���id = b.���id) " & IIf(strSql = "", "", "Minus (" & Mid(strSql, 8) & ") ")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "MsgDis", gclsPros.����ID, gclsPros.��ҳID, str����IDs, str���IDs)
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ���˴�Ⱦ��������ݷ����˸ı�,���޸Ĵ�Ⱦ�����濨��", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub


