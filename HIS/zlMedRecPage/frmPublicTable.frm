VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPublicTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ҽ���"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11640
   Icon            =   "frmPublicTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   11640
   Begin VSFlex8Ctl.VSFlexGrid vsTable 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _cx             =   20558
      _cy             =   5530
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPublicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mintType As Integer
Public mfrmParent As Form
Public mlng����ID As Long, mlng��ҳID As Long
Public mlngLeft As Long, mlngTop As Long, mlngHeight As Long
Public mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, frmParent As Form, ByVal x As Long, ByVal Y As Long, ByVal lngHeight As Long) As Boolean
'���أ�ShowMe= ��ȷ������ȡ��
'����:intType 1-��ҽ��� 2-��ҽ��� 3-������¼
'     frmParent ������
    mintType = intType
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngLeft = x
    mlngTop = Y
    mlngHeight = lngHeight
    Set mfrmParent = frmParent
    Set mrsTmp = LoadTableData(mintType)
    If mlng����ID = 0 And mlng��ҳID = 0 Then
        MsgBox "����û��ѡ���ˣ������ڼ�¼��", vbInformation, gstrSysName
        Exit Function
    Else
        If Not mrsTmp Is Nothing Then
            If mrsTmp.RecordCount < 1 Then
                If mintType = 1 Then
                    MsgBox "û���ҵ���¼��Դ����סԺ��ҳ��ҽ����ҽ��ϼ�¼��", vbInformation, gstrSysName
                ElseIf mintType = 2 Then
                    MsgBox "û���ҵ���¼��Դ����סԺ��ҳ��ҽ����ҽ��ϼ�¼��", vbInformation, gstrSysName
                ElseIf mintType = 3 Then
                    MsgBox "û���ҵ���¼��Դ����סԺ��ҳ��ҽ��������¼��¼��", vbInformation, gstrSysName
                End If
                Exit Function
            Else
                Show 0, frmParent
            End If
        Else
            MsgBox "�ò��˲����ڼ�¼��Դ����סԺ��ҳ" & IIf(mintType = 1 Or mintType = 1 = 2, "��ϼ�¼", "������¼"), , vbInformation, gstrSysName
            Exit Function
        End If
    End If
    ShowMe = True
End Function

Private Function InitTable(ByVal intType As Integer) As Boolean
    Dim strHead As String
    Dim strRow As String
On Error GoTo errH
    Select Case intType
        Case 1
            strHead = "����������ÿ�,1250,4;����;��ϱ���,900,4;�������,3200,1;��ҽ֤��;����ʱ��;��ע,1200,1;��Ժ����,850,1;��Ժ���,850,1;ICD����,800,1;δ��,350,4;����,350,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
            strRow = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���XY & ";" & _
                                    DI_������� & ",Ժ�ڸ�Ⱦ," & DI_��Ϸ��� & "," & DT_Ժ�ڸ�Ⱦ & ";" & _
                                    DI_������� & ", �� �� ֢ ," & DI_��Ϸ��� & "," & DT_����֢ & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_������� & ";" & _
                                    DI_������� & ",�����ж�," & DI_��Ϸ��� & "," & DT_�����ж���
        Case 2
            strHead = "����������ÿ�,1250,4;����;��ϱ���,900,4;�������,3000,1;��ҽ֤��,1500,1;����ʱ��;��ע,1100,1;��Ժ����,850,1;��Ժ���,850,1;ICD����;δ��;����,350,4;" & _
                                        ",270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;�Ƿ���;��Ч����;������Ϣ;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
            strRow = DI_������� & ",�ţ����������," & DI_��Ϸ��� & "," & DT_�������ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",��Ժ���," & DI_��Ϸ��� & "," & DT_��Ժ���ZY & ";" & _
                                    DI_������� & ",�������," & DI_��Ϸ��� & "," & DT_��Ժ���ZY
        Case 3
            If gclsPros.MedPageSandard = ST_��������׼ Then
                strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "������ʼʱ��,1850,4;��������ʱ��,1850,4", "��������������,1850,4;��������ʱ��") & ";��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������������,1500,1;��������������,2800,1;�ٴ�����,850,4,11;����,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                                "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;��������,850,1;����ҽʦ,850,1;�п����ϵȼ�,1400,1;�пڲ�λ,850,1;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                                "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������,1600,4,11;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                                "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
            ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "������ʼʱ��,1850,4;��������ʱ��,1850,4", "��������������,1850,4;��������ʱ��") & ";��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������������,1500,1;��������������,2800,1;�ٴ�����,850,4,11;����,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                                "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;��������,850,1;����ҽʦ,850,1;�п����ϵȼ�,1400,1;�пڲ�λ;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                                "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                                "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
            ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                strHead = ",300,4;" & "��ʼ����,1850,4;��������,1850,4;��ǰԤ���Կ�����ҩʱ��,2150,4;�������,875,1;׼������,850,7;��������,1500,1;��������,2800,1;�ٴ�����,850,4,11;����ҽʦ,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                                "����ʼʱ��,1550,4;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;�����ּ�,850,1;����ҽʦ,850,1;�п�/����,1400,1;�пڲ�λ,850,1;�ط������Ҽƻ�,1400,4,11;�ط�������Ŀ��,1400,1;�пڸ�Ⱦ,850,4,11;����֢,720,4,11;" & _
                                "��ǰ0.5-2СʱԤ���ÿ���ҩ;�������Χ����Ԥ���ÿ���ҩ����;��Ԥ�ڵĶ�������;������֢;������������;��������֢;�����Ѫ��Ѫ��;�����˿��ѿ�;�������Ѫ˨;��������/��л����;�������˥��;" & _
                                "�����˨��;�����Ѫ֢;�����Źؽڹ���;��������ID;������ĿID;����ID;��������;������Դ"
            ElseIf gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                strHead = ",300,4;" & "��������,1850,4;��������;��ǰԤ���Կ�����ҩʱ��;�������,875,1;׼������;��������,1500,1;��������,2800,1;�ٴ�����,850,4,11;����ҽʦ,850,1;������ʿ,850,1;�ڢ�����,850,1;�ڢ�����,850,1;" & _
                                "����ʼʱ��;����ʽ,850,1;ASA�ּ�,850,1;NNIS�ּ�,850,1;�����ּ�,850,1;����ҽʦ,850,1;�п�/����,1400,1;�пڲ�λ;�ط������Ҽƻ�;�ط�������Ŀ��;�пڸ�Ⱦ;����֢;" & _
                                "��ǰ0.5-2СʱԤ���ÿ���ҩ,2400,4,11;�������Χ����Ԥ���ÿ���ҩ����,2850,7;��Ԥ�ڵĶ�������,1600,4,11;������֢,1000,4,11;������������,1200,4,11;��������֢,1000,4,11;" & _
                                "�����Ѫ��Ѫ��,1450,4,11;�����˿��ѿ�,1200,4,11;�������Ѫ˨,1450,4,11;��������/��л����,1700,4,11;�������˥��,1200,4,11;�����˨��,1000,4,11;�����Ѫ֢,1000,4,11;" & _
                                "�����Źؽڹ���,1450,4,11;��������ID;������ĿID;����ID;��������;������Դ"
            End If
    End Select
    If Not setTableType(intType, strHead, strRow) Then Exit Function
    InitTable = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function setTableType(ByVal intType As Integer, Optional ByVal strHead As String, Optional ByVal strRow As String) As Boolean
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
On Error GoTo errH
    Select Case intType
        Case 1
            Set vsTmp = vsTable
            Call Grid.Init(vsTable, strHead, strRow, 1, 1)
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
        Case 2
            Set vsTmp = vsTable
            Call Grid.Init(vsTable, strHead, strRow, 1, 1)
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
        Case 3
            Set vsTmp = vsTable
            Call Grid.Init(vsTmp, strHead)
            With vsTmp
                .Font.Size = 9
                If gclsPros.FuncType <> f���Ӳ��� Then
                    .ColComboList(PI_�������) = " |����|����|����"
                    .ColComboList(PI_ASA�ּ�) = " |P1|P2|P3|P4|P5|P6"
                    .ColComboList(PI_NNIS�ּ�) = " |NNIS0��|NNIS1��|NNIS2��|NNIS3��"
                    .ColComboList(PI_��������) = " |��|һ������|��������|��������|�ļ�����"
                    '�п�����
                    strSql = "Select Rownum As ID, To_Number(����) As ����, ���� ����, ����, 0 ȱʡ From �����п����� Order By ����"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����п�����")
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
    End Select
    setTableType = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadTableData(ByVal intType As Integer) As ADODB.Recordset
    Dim strSql As String, strSQLTmp As String, strDiagType As String, strSQLJudge As String
    Dim rsTmp As New ADODB.Recordset
    Dim int��¼��Դ As Integer
On Error GoTo errH
    Select Case intType
        Case 1, 2
            If gclsPros.FuncType = f������ҳ Then
               int��¼��Դ = 3
               strSQLJudge = "Select 1 From ������ϼ�¼ Where ����id = [1] And ��ҳid =[2] And ��¼��Դ = [3] And Rownum < 2"
               Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "��ҳ��Դ����ж�", mlng����ID, mlng��ҳID, int��¼��Դ)
               If rsTmp.RecordCount > 0 Then
                   strDiagType = " And A.��¼��Դ =[3] "
               Else
                   strDiagType = ""
                   Set LoadTableData = rsTmp
                   Exit Function
               End If
            End If
            If intType = 1 Then
                strDiagType = strDiagType & " And A.������� IN(1,2,3,5,6,7,10,21) "
            Else
                strDiagType = strDiagType & " And A.������� IN(1,2,3,5,6,7,10,11,12,13,21) "
            End If
            strSql = "Select A.��ע, A.Id, A.����id, A.��ҳid, A.ҽ��id, A.��¼��Դ, A.��ϴ���, Nvl(A.�������,1) �������, A.�������, A.��Ժ����, A.����id, A.���id, A.֤��id,B.���� ��������,C.���� �������,D.���� ֤������," & vbNewLine & _
                "       A.�������, A.��Ժ���, A.�Ƿ�δ��, A.�Ƿ�����, A.����ʱ��, B.���� As ��������,B.��� As �������, B.����, C.���� As ��ϱ���, D.���� As ֤�����," & vbNewLine & _
                IIf(gclsPros.FuncType = f���Ӳ���, " Null ҽ��id", " (Select F_List2str(Cast(Collect(C.ҽ��id || '') As T_Strlist)) ҽ��id" & vbNewLine & _
                "         From �������ҽ�� C,����ҽ����¼ F " & vbNewLine & _
                "         Where C.ҽ��ID = F.ID and C.���id = A.Id and nvl(F.�������,0) = 0) As ҽ��id") & ",B.�Ա�����, B.��Ч����, B.����, B.����, E.Id As ����, E.�Ƿ���,Null ����ID,A.��¼����,A.��¼�� " & vbNewLine & _
                "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C, ��������Ŀ¼ D,����������� E" & vbNewLine & _
                "Where A.����id = B.Id(+) And A.���id = C.Id(+) And A.֤��id = D.Id(+)  And  B.����id = E.Id(+)" & strDiagType & "And A.ȡ��ʱ�� Is Null And A.������� Is Not Null And ����id = [1] And ��ҳid =[2]" & vbNewLine & _
                "Order By A.�������, A.��¼��Դ Desc, A.��ϴ���, Nvl(A.�������,1), A.Id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ҳ���", mlng����ID, mlng��ҳID, int��¼��Դ)
            Set LoadTableData = rsTmp
        Case 3
            strSql = "Select A.Id, A.����id, A.��ҳid, A.�������, A.��¼��Դ, A.��������, A.������ʼʱ��, A.��������ʱ��, Nvl(B.����, C.����) As ��������, A.�������� ��������," & vbNewLine & _
                "       Nvl(B.����, C.����) ����ԭ��, A.����ҽʦ, A.������ʿ, A.��һ����, A.�ڶ�����, A.����ҽʦ, A.׼������, A.������ҩʱ��, A.������ҩ����, A.����ʼʱ��, A.�ط�Ŀ��," & vbNewLine & _
                "       A.�пڲ�λ, A.��������, Decode(A.Asa�ּ�, 'I��', 'P1', 'II��', 'P2', 'III��', 'P3', 'IV��', 'P4', 'V��', 'P5', A.Asa�ּ�) Asa�ּ�, A.Nnis�ּ�, Decode(A.��������, 1, 'һ������', 2, '��������', 3, '��������', 4, '�ļ�����',9, '��', ' ') As ��������, A.�п�," & vbNewLine & _
                "       A.����, A.�ٴ�����, A.��ǰ������ҩ, A.��Ԥ�ڵĶ�������, A.������֢, A.������������, A.��������֢, A.�����Ѫ��Ѫ��, A.�����˿��ѿ�, A.�������Ѫ˨, A.���������л����," & vbNewLine & _
                "       A.�������˥��, A.�����˨��, A.�����Ѫ֢, A.�����Źؽڹ���, A.�ط��ƻ�, A.�пڸ�Ⱦ, A.����֢, A.��������id, A.������Ŀid, A.����ʽ ����id, D.���� ����ʽ, A.��¼����," & vbNewLine & _
                "       A.��¼��, A.ȡ��ʱ��, A.ȡ����, Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) ԭ�������� " & vbNewLine & _
                "From ���������¼ A, ��������Ŀ¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D" & vbNewLine & _
                "Where C.Id(+) = A.������Ŀid And A.��������id = B.Id(+) And A.����ʽ = D.Id(+) And ����id = [1] And ��ҳid = [2] And" & vbNewLine & _
                "      (��¼��Դ <> 1 Or" & vbNewLine & _
                "       (��¼��Դ = 1 And ȡ��ʱ�� Is Null And" & vbNewLine & _
                "       ��¼���� =" & vbNewLine & _
                "       (Select Max(��¼����) From ���������¼ Where ����id =[1] And ��ҳid = [2] And ȡ��ʱ�� Is Null)))" & vbNewLine & _
                "Order By Nvl(A.��������,999),A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", mlng����ID, mlng��ҳID)
            rsTmp.Filter = "��¼��Դ=3"
            Set LoadTableData = rsTmp
        End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsDiagData(ByRef vsTable As VSFlexGrid, ByVal rsInput As ADODB.Recordset, ByVal strDiagType As String)
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
    
On Error GoTo errH
    blnGet���� = gclsPros.GetExtraCode
    arrTmp = Split(strDiagType, ",")
    bln��ҽ = mintType = 1
    With vsTable
        For i = LBound(arrTmp) To UBound(arrTmp)
            Call FilterDiagByType(rsInput, Val(arrTmp(i)), -1) '�������
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
                    If Not (IsNull(rsInput!���ID) And IsNull(rsInput!����id)) Then
                        .Cell(flexcpData, LngRow, DI_�������) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!������� & "")
                    Else
                        .Cell(flexcpData, LngRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
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
    End With
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsOPSData(ByRef vsOPSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset)
'���ܣ����ز����������ݲ�����
'������vsOPSInput=��Ҫ���ز���������Ϣ�ı��
'      rsInput=����������Ϣ��¼��
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    Dim lngOrder As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    With vsOPSInput
        '���ݼ���
        If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Function
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
    End With
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    Dim lngScrH  As Long
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    If mlngTop + Me.Height > lngScrH Then
        Me.Top = mlngTop - Me.Height - 300
    Else
        Me.Top = mlngHeight + 1000
    End If
    Me.Left = mlngLeft
    If Not InitTable(mintType) Then Exit Sub
    Select Case mintType
        Case 1, 2
            mrsTmp.Filter = "��¼��Դ=3"
            Call LoadVsDiagData(vsTable, mrsTmp, IIf(mintType = 1, "1,2,3,5,6,7,10", "11,12,13"))
        Case 3
            Call LoadVsOPSData(vsTable, mrsTmp)
    End Select
    If mintType = 1 Then
        Me.Caption = "ҽ����ҽ���"
    ElseIf mintType = 2 Then
        Me.Caption = "ҽ����ҽ���"
    ElseIf mintType = 3 Then
        Me.Caption = "ҽ��������¼"
    End If
    Exit Sub
End Sub



